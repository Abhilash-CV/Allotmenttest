#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CEE Kerala CAP allotment engine - rewrite of allotment_process.php

INPUT  (Excel, first sheet of each)
    --options     optionentry.xlsx     RollNo, OPNO, Optn, ValidOption
    --seats       seatcategory.xlsx    CounselGroup, CollegeType, CollegeCode, CourseCode, Category, Seat
    --candidates  candidate.xlsx       RollNo, MRank, URank, ARank, Category, Others, Special1-4, Keralite,
                                       NRI, MCat1, MCat2, NCC, NCCMBBSRank, NCCMRank, PI_MMRank, PT_MMRank,
                                       PI_MRank, PT_MRank, LSMRank, (FW)
    --restrict    restrict_rank.xlsx   OPTIONAL: CollegeCode, CourseCode, Grp, Category, MaxRank
OUTPUT
    one workbook: Allotment, Seats, Rounds

Option string  Optn = [0]group [1]college type [2:4]course [4:7]college [7]fee code   e.g. MGAACOAG
Allot string   = Optn[:7] + quota(4)   SMSM open merit | SMEZ category candidate floated INTO SM
                                       FLEZ back-float (category seat made from an SM seat) | EZEZ category seat

What is the same as the PHP
  * passes per round: M, U, A, then NCC-MBBS, NCC-other, PI/PT (MBBS and other), LS
  * every round resets seats and re-runs all passes; a candidate may only look at options <= the option held
    from the previous pass/round when that hold came from a special quota (OPNO_O / Allot_O logic)
  * rounds repeat until no candidate's option changes ("jump" == 0)
  * govt-college float logic: state-wide SM pool per float group, float-in / back-float
What is different (on purpose)
  * no database, no global arrays, no per-row UPDATEs
  * float-in / float-out are applied symmetrically (cap = total + in - out) so SM availability also
    respects seats floated out
  * when a candidate is moved to a new seat inside a round, the old seat is released
    (PHP kept it consumed).  Set RELEASE_ON_OVERWRITE = False to reproduce the PHP behaviour
  * float group is (CounselGroup, Course) for CollegeType G unless seatcategory has a FloatGroup column
  * Last rank per seat bucket is computed from the final result
"""
import argparse
import io
import contextlib
import streamlit as st
import sys
from collections import defaultdict, namedtuple

import pandas as pd

# ----------------------------------------------------------------------------
# switches
# ----------------------------------------------------------------------------
MAX_ROUNDS = 50
SINGLE_ROUND = False          # PHP: defvalues SlNo=12 == 'N'  -> stop after round 1
RELEASE_ON_OVERWRITE = True
USE_FW = False                # PHP never loaded FW for the M/A/U query, so it was dead code

MEDI = {'MM', 'MD', 'BA', 'BH', 'BS', 'BU'}
HELD = {'PIPI', 'PTPT', 'CCCC'}

MAIN_PASSES = [('M', 'MRANK'), ('U', 'URANK'), ('A', 'ARANK')]

SPECIALS = [
    dict(name='NCC-MBBS',  rank='NCCMBBSRANK', need_ncc=True, quota='CC', opt_in=MEDI, prev_group='M'),
    dict(name='NCC-OTHER', rank='NCCMRANK',    need_ncc=True, quota='CC', opt_out=MEDI, prev_group='M'),
    dict(name='PI-MBBS',   rank='PI_MMRANK',   quota='PI', opt_in=MEDI),
    dict(name='PT-MBBS',   rank='PT_MMRANK',   quota='PT', opt_in=MEDI),
    dict(name='PI-OTHER',  rank='PI_MRANK',    quota='PI', opt_out=MEDI, prev_extra_out={'BP'}),
    dict(name='PT-OTHER',  rank='PT_MRANK',    quota='PT', opt_out=MEDI, prev_extra_out={'BP'}),
    dict(name='LS',        rank='LSMRANK',     quota='LS', opt_group=None, prev_group='M'),
]
RANK_COLS = [r for _, r in MAIN_PASSES] + [s['rank'] for s in SPECIALS]


def course_ok(g, course):
    if g == 'M':
        return course not in ('BA', 'BH', 'BS', 'BU')
    if g == 'A':
        return course == 'BA'
    return course in ('BH', 'BS', 'BU')     # 'U'


# ----------------------------------------------------------------------------
# Excel helpers
# ----------------------------------------------------------------------------
def clean(v):
    v = '' if v is None else str(v).strip()
    return '' if v.upper() in ('NULL', 'NONE') else v


def num(v):
    try:
        return float(clean(v) or 0)
    except ValueError:
        return 0.0


def rollkey(v):
    v = clean(v)
    return v[:-2] if v.endswith('.0') else v


def load(path):
    # keep_default_na=False so the category code "NA" is not turned into NaN
    df = pd.read_excel(path, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def find(df, name, required=True):
    """Exact column, or a header truncated by the export (COLLEGECOD -> COLLEGECODE)."""
    if name in df.columns:
        return name
    for c in df.columns:
        if len(c) >= 6 and (name.startswith(c) or c.startswith(name)):
            return c
    if required:
        raise KeyError('column %s not found; have %s' % (name, list(df.columns)))
    return None


Opt = namedtuple('Opt', 'opno optn grp ctype course college fee')


def load_options(path):
    df = load(path)
    rc, oc, tc = find(df, 'ROLLNO'), find(df, 'OPNO'), find(df, 'OPTN')
    vc = find(df, 'VALIDOPTION', False)
    out = defaultdict(list)
    for r in df.to_dict('records'):
        if vc and clean(r[vc]).upper() != 'Y':
            continue
        s = clean(r[tc]).upper()
        if len(s) < 7:
            continue
        out[rollkey(r[rc])].append(Opt(int(num(r[oc])), s, s[0], s[1], s[2:4], s[4:7], s[7:8]))
    for v in out.values():
        v.sort(key=lambda o: o.opno)
    return out


def load_seats(path):
    df = load(path)
    cg, ct, co = find(df, 'COUNSELGROUP'), find(df, 'COLLEGETYPE'), find(df, 'COLLEGECODE')
    cr, ca, cs = find(df, 'COURSECODE'), find(df, 'CATEGORY'), find(df, 'SEAT')
    fgc = find(df, 'FLOATGROUP', False)
    tot, gtot, fgmap = defaultdict(int), defaultdict(int), {}
    for r in df.to_dict('records'):
        raw = [clean(r[x]) for x in (cg, ct, co, cr, ca)]
        if not all(raw) or any(v != v.upper() for v in raw):      # blank / lower-case "type" row
            continue
        grp, typ, col, crs, cat = raw
        seat = int(num(r[cs]))
        tot[(col, crs, cat)] += seat
        if typ == 'G':
            fg = clean(r[fgc]) if fgc else (grp, crs)
            fgmap[(col, crs)] = fg
            gtot[(fg, cat)] += seat
    return dict(tot), dict(gtot), fgmap


def load_restrict(path):
    out = {}
    if not path:
        return out
    df = load(path)
    a, b, g, c, m = (find(df, n) for n in ('COLLEGECODE', 'COURSECODE', 'GRP', 'CATEGORY', 'MAXRANK'))
    for r in df.to_dict('records'):
        out[(clean(r[a]).upper(), clean(r[b]).upper(), clean(r[g]).upper(), clean(r[c]).upper())] = num(r[m])
    return out


class Cand(object):
    __slots__ = ('roll', 'name', 'ranks', 'cat', 'others', 'sp', 'keralite', 'nri', 'mcat1', 'mcat2',
                 'fw', 'ncc', 'opts', 'opno', 'allot', 'rank', 'prev_opno', 'prev_allot', 'undo')


def load_candidates(path, options):
    df = load(path)
    rc = find(df, 'ROLLNO')
    cols = {c: find(df, c, False) for c in RANK_COLS}
    g = lambda r, n: clean(r[n]).upper() if n in r else ''
    out = []
    for r in df.to_dict('records'):
        c = Cand()
        c.roll = rollkey(r[rc])
        c.name = clean(r.get('NAME', ''))
        c.ranks = {k: num(r[v]) for k, v in cols.items() if v}
        if not any(c.ranks.values()):
            continue
        c.cat, c.others = g(r, 'CATEGORY'), g(r, 'OTHERS')
        c.sp = [g(r, 'SPECIAL%d' % i) for i in (1, 2, 3, 4)]
        c.keralite, c.nri = g(r, 'KERALITE'), g(r, 'NRI')
        c.mcat1, c.mcat2 = g(r, 'MCAT1'), g(r, 'MCAT2')
        c.fw, c.ncc = g(r, 'FW'), g(r, 'NCC')
        c.opts = options.get(c.roll, [])
        c.opno, c.allot, c.rank, c.prev_opno, c.prev_allot, c.undo = 0, '', 0, 0, '', []
        out.append(c)
    return out


# ----------------------------------------------------------------------------
# seat book
# ----------------------------------------------------------------------------
class Book(object):
    def __init__(self, tot, gtot, fgmap):
        self.tot, self.gtot, self.fg = tot, gtot, fgmap
        self.reset()

    def reset(self):
        self.alloc, self.fin, self.fout = defaultdict(int), defaultdict(int), defaultdict(int)
        self.gall, self.gfl = defaultdict(int), defaultdict(int)

    def cap(self, k):
        return self.tot.get(k, 0) + self.fin.get(k, 0) - self.fout.get(k, 0)

    def avail(self, col, crs, cat):
        k = (col, crs, cat)
        return self.cap(k) - self.alloc.get(k, 0)

    def sm_left(self, fg):
        return self.gtot.get((fg, 'SM'), 0) - self.gall.get((fg, 'SM'), 0)

    def apply(self, ops, sign=1):
        for name, key, d in ops:
            getattr(self, name)[key] += sign * d


# ----------------------------------------------------------------------------
# engine
# ----------------------------------------------------------------------------
class Engine(object):
    def __init__(self, book, cands, restrict):
        self.book, self.cands, self.restrict = book, cands, restrict
        self.jump = 0
        self.stats = []

    # ---- round / pass bookkeeping --------------------------------------
    def start_round(self):
        for c in self.cands:
            if c.allot == '':                         # lost its seat last round -> forget the hold
                c.prev_allot, c.prev_opno = '', 0
            c.opno, c.allot, c.rank, c.undo = 0, '', 0, []
        self.book.reset()
        self.jump = 0

    def end_pass(self):                               # update candidate set OPNO_O=OPNO,Allot_O=Allot where OPNO!=0
        for c in self.cands:
            if c.opno != 0:
                c.prev_opno, c.prev_allot = c.opno, c.allot

    def commit(self, c, o, ops, code, rank):
        if RELEASE_ON_OVERWRITE and c.undo:
            self.book.apply(c.undo, -1)
        self.book.apply(ops)
        c.undo = ops
        if o.opno != c.prev_opno:
            self.jump += 1
        c.opno, c.allot, c.rank = o.opno, code, rank

    # ---- main passes (M / U / A) ---------------------------------------
    @staticmethod
    def main_opts(c, g):
        po, pa = c.prev_opno, c.prev_allot
        full = po == 0 or (pa[:1] == 'M' and course_ok(g, pa[2:4]) and pa[7:11] not in HELD)
        opts = c.opts if full else [o for o in c.opts if o.opno <= po]
        return [o for o in opts if o.grp == 'M' and course_ok(g, o.course)]

    def lim(self, g, col, crs, cat, rank):
        mr = self.restrict.get((col, crs, g, cat))
        return mr is not None and rank > mr

    def try_option(self, c, o, rank, g):
        """Return (ops, allot_code) for the first seat this candidate can take on this option, else None."""
        b = self.book
        col, crs, o7 = o.college, o.course, o.optn[:7]
        cat = '' if c.cat in ('', 'NA') else c.cat

        def has(q):
            return q not in ('', 'NA') and b.avail(col, crs, q) > 0

        def normal(q):
            return [('alloc', (col, crs, q), 1)], o7 + q + q

        # -- government colleges, non-management fee: SM pool + floating
        if o.ctype == 'G' and o.fee != 'M':
            fg = b.fg.get((col, crs))
            sm_left = b.sm_left(fg)
            if self.lim(g, col, crs, 'SM', rank):
                sm_left = 0
            can_float = cat != '' and not (o.grp == 'B' and cat == 'EW')
            if USE_FW and c.fw == 'FW' and has('FW'):
                return normal('FW')
            if sm_left > 0:
                if b.avail(col, crs, 'SM') > 0:                                   # open merit
                    return ([('alloc', (col, crs, 'SM'), 1), ('gall', (fg, 'SM'), 1)], o7 + 'SMSM')
                if can_float and has(cat):                                        # float IN to SM
                    return ([('alloc', (col, crs, 'SM'), 1), ('fin', (col, crs, 'SM'), 1),
                             ('fout', (col, crs, cat), 1), ('gall', (fg, 'SM'), 1), ('gfl', (fg, cat), 1)],
                            o7 + 'SM' + cat)
            elif can_float:
                chk = b.gtot.get((fg, cat), 0) - b.gall.get((fg, cat), 0)
                if b.avail(col, crs, cat) > 0 and chk > 0:                        # own category seat
                    if self.lim(g, col, crs, cat, rank):
                        return None
                    return ([('alloc', (col, crs, cat), 1), ('gall', (fg, cat), 1)], o7 + cat + cat)
                if chk > 0 and b.gfl.get((fg, cat), 0) > 0 and b.avail(col, crs, 'SM') > 0:   # back-float
                    if self.lim(g, col, crs, cat, rank):
                        return None
                    return ([('fout', (col, crs, 'SM'), 1), ('alloc', (col, crs, cat), 1),
                             ('fin', (col, crs, cat), 1), ('gall', (fg, cat), 1), ('gfl', (fg, cat), -1)],
                            o7 + 'FL' + cat)
            if o.grp == 'B' and cat == 'EW' and has('EW'):
                return normal('EW')
            if cat == 'SC' and has('PC'):
                return normal('PC')
            for q in [c.others] + c.sp:
                if has(q):
                    return normal(q)

        # -- government, management-fee seats
        if o.ctype == 'G' and o.fee == 'M' and c.keralite == 'Y':
            if c.sp[2] == 'PM' and has('PM'):
                return normal('PM')
            if has('MG'):
                return normal('MG')

        # -- self-financing / NRI-type colleges, general fee
        if o.ctype in ('S', 'N') and o.fee == 'G':
            if has('AM'):
                return normal('AM')
            avl = b.avail(col, crs, 'SM')
            if o.grp == 'M' and c.keralite == '2' and crs not in ('BS', 'BU'):
                avl = 0
            if avl > 0:
                return normal('SM')
            for q in [cat, c.others, c.sp[0], c.sp[1]] + [x for x in c.sp[2:] if x != 'CB']:
                if has(q):
                    return normal(q)

        if o.fee == 'Z' and 'CB' in (c.sp[2], c.sp[3]) and has('CB'):
            return normal('CB')

        if o.ctype == 'S' and o.fee == 'M':
            if has('FD'):
                return normal('FD')
            if c.sp[2] in ('CB', 'PM') and has(c.sp[2]):
                return normal(c.sp[2])
            if has('MG'):
                return normal('MG')

        if c.mcat2 and o.ctype == 'S' and o.fee == 'Y':                          # medical minority
            for q in [c.mcat2[i:i + 2] for i in range(0, len(c.mcat2), 2)]:
                if has(q):
                    return normal(q)

        if o.fee == 'R' and c.nri in ('Y', 'P') and c.mcat1:                    # NRI
            for q in [c.mcat1[i:i + 2] for i in range(0, len(c.mcat1), 2)]:
                if has(q):
                    return normal(q)
        return None

    def run_main(self, g, rankcol):
        order = sorted((c for c in self.cands if c.ranks.get(rankcol, 0) != 0),
                       key=lambda c: (c.ranks[rankcol], c.roll))
        for c in order:
            rank = c.ranks[rankcol]
            for o in self.main_opts(c, g):
                hit = self.try_option(c, o, rank, g)
                if hit:
                    self.commit(c, o, hit[0], hit[1], rank)
                    break
        self.end_pass()

    # ---- special quotas (NCC / PI / PT / LS) ----------------------------
    @staticmethod
    def special_opts(c, sp):
        po, pa = c.prev_opno, c.prev_allot
        full = po == 0
        if not full and pa[7:11] == sp['quota'] * 2 and sp.get('prev_group') in (None, pa[:1]):
            pc = pa[2:4]
            if 'opt_in' in sp:
                full = pc in sp['opt_in']
            elif 'opt_out' in sp:
                full = pc not in sp['opt_out'] and pc not in sp.get('prev_extra_out', ())
            else:
                full = True
        opts = c.opts if full else [o for o in c.opts if o.opno <= po]
        gsel = sp.get('opt_group', 'M')
        return [o for o in opts
                if (not gsel or o.grp == gsel)
                and ('opt_in' not in sp or o.course in sp['opt_in'])
                and ('opt_out' not in sp or o.course not in sp['opt_out'])]

    def run_special(self, sp):
        key, q = sp['rank'], sp['quota']
        todo = [c for c in self.cands
                if c.ranks.get(key, 0) != 0 and (not sp.get('need_ncc') or c.ncc == 'CC')]
        todo.sort(key=lambda c: (c.ranks[key], c.roll))
        for c in todo:
            for o in self.special_opts(c, sp):
                if self.book.avail(o.college, o.course, q) > 0:
                    # already holds this very option through another quota -> keep it
                    if not (c.prev_opno == o.opno and c.allot != ''):
                        self.commit(c, o, [('alloc', (o.college, o.course, q), 1)],
                                    o.optn[:7] + q + q, c.ranks[key])
                    break
        self.end_pass()

    # ---- driver --------------------------------------------------------
    def run(self):
        converged = False
        for rnd in range(1, MAX_ROUNDS + 1):
            self.start_round()
            for g, rankcol in MAIN_PASSES:
                self.run_main(g, rankcol)
            for sp in SPECIALS:
                self.run_special(sp)
            n = sum(1 for c in self.cands if c.allot)
            self.stats.append(dict(ROUND=rnd, JUMPS=self.jump, ALLOTTED=n))
            print('round %d: jumps=%d allotted=%d' % (rnd, self.jump, n))
            if SINGLE_ROUND or self.jump == 0:
                converged = True
                break
        if not converged:
            print('WARNING: no convergence in %d rounds' % MAX_ROUNDS, file=sys.stderr)


# ----------------------------------------------------------------------------
# output
# ----------------------------------------------------------------------------
def bucket(code):
    """physical seat bucket a candidate occupies"""
    q = code[7:11]
    if q[:2] == 'SM':
        return 'SM'
    if q[:2] == 'FL':
        return q[2:]
    return q[:2]


def write_report(eng, path):
    rows, last = [], {}
    for c in sorted(eng.cands, key=lambda c: (c.ranks.get('MRANK', 0) or 1e12, c.roll)):
        a = c.allot
        rows.append(dict(ROLLNO=c.roll, NAME=c.name, CATEGORY=c.cat, MRANK=c.ranks.get('MRANK', 0),
                         ALLOT=a, OPNO=c.opno or '', GROUP=a[:1], COLLEGETYPE=a[1:2], COURSE=a[2:4],
                         COLLEGE=a[4:7], QUOTA=a[7:11], STATUS='ALLOTTED' if a else 'NOT ALLOTTED'))
        if a:
            k = (a[4:7], a[2:4], bucket(a))
            last[k] = max(last.get(k, 0), c.rank)
    b = eng.book
    keys = set(b.tot) | set(b.alloc) | set(b.fin) | set(b.fout)
    seats = []
    for k in sorted(keys):
        seats.append(dict(COLLEGE=k[0], COURSE=k[1], CATEGORY=k[2], SEAT=b.tot.get(k, 0),
                          ALLOTTED=b.alloc.get(k, 0), FLOAT_IN=b.fin.get(k, 0), FLOAT_OUT=b.fout.get(k, 0),
                          VACANT=b.cap(k) - b.alloc.get(k, 0), LAST_RANK=last.get(k, 0)))
        if b.alloc.get(k, 0) > b.cap(k):
            print('WARNING over-allocated', k, file=sys.stderr)
    with pd.ExcelWriter(path, engine='openpyxl') as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name='Allotment', index=False)
        pd.DataFrame(seats).to_excel(xw, sheet_name='Seats', index=False)
        pd.DataFrame(eng.stats).to_excel(xw, sheet_name='Rounds', index=False)
    print('written', path)

def main():
    global SINGLE_ROUND, RELEASE_ON_OVERWRITE

    st.set_page_config(page_title="CEE Kerala CAP Allotment Engine", layout="wide")
    st.title("CEE Kerala CAP Allotment Engine")
    st.caption("Upload the input Excel files, run the allotment process, and download the result workbook.")

    with st.expander("Required input formats", expanded=True):
        st.markdown("""
        - **Options**: `RollNo`, `OPNO`, `Optn`, `ValidOption` (optional)
        - **Seats**: `CounselGroup`, `CollegeType`, `CollegeCode`, `CourseCode`, `Category`, `Seat`
        - **Candidates**: `RollNo`, rank columns such as `MRank`, `URank`, `ARank`, category/special-quota fields
        - **Rank restrictions** (optional): `CollegeCode`, `CourseCode`, `Grp`, `Category`, `MaxRank`

        Upload `.xlsx` or `.xls` files. The engine reads the first worksheet in each workbook.
        """)

    col1, col2 = st.columns(2)
    with col1:
        options_file = st.file_uploader("Options file", type=["xlsx", "xls"], key="options")
        seats_file = st.file_uploader("Seat matrix file", type=["xlsx", "xls"], key="seats")
    with col2:
        candidates_file = st.file_uploader("Candidates file", type=["xlsx", "xls"], key="candidates")
        restrict_file = st.file_uploader("Rank restrictions (optional)", type=["xlsx", "xls"], key="restrict")

    st.subheader("Run settings")
    c1, c2 = st.columns(2)
    with c1:
        single_round = st.checkbox("Single round only", value=False,
                                   help="Stop after the first round instead of repeating until no option changes.")
    with c2:
        legacy_leak = st.checkbox("Legacy seat behaviour",
                                  value=False,
                                  help="When selected, the old seat remains consumed when a candidate moves.")

    output_name = st.text_input("Output filename", value="allotment_result.xlsx")
    if not output_name.strip():
        output_name = "allotment_result.xlsx"
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"

    if st.button("Run allotment", type="primary", use_container_width=True):
        missing = []
        if options_file is None:
            missing.append("Options file")
        if seats_file is None:
            missing.append("Seat matrix file")
        if candidates_file is None:
            missing.append("Candidates file")

        if missing:
            st.error("Please upload the required file(s): " + ", ".join(missing))
            st.stop()

        try:
            SINGLE_ROUND = single_round
            RELEASE_ON_OVERWRITE = not legacy_leak

            with st.spinner("Running allotment engine..."):
                log_buffer = io.StringIO()
                err_buffer = io.StringIO()
                with contextlib.redirect_stdout(log_buffer), contextlib.redirect_stderr(err_buffer):
                    options = load_options(options_file)
                    tot, gtot, fgmap = load_seats(seats_file)
                    cands = load_candidates(candidates_file, options)
                    eng = Engine(Book(tot, gtot, fgmap), cands, load_restrict(restrict_file))
                    eng.run()

                    result_buffer = io.BytesIO()
                    write_report(eng, result_buffer)
                    result_buffer.seek(0)

            st.success("Allotment processing completed.")
            m1, m2, m3 = st.columns(3)
            m1.metric("Candidates loaded", len(cands))
            m2.metric("Seat buckets", len(tot))
            allotted_count = sum(1 for candidate in cands if candidate.allot)
            m3.metric("Allotted candidates", allotted_count)

            logs = log_buffer.getvalue().strip()
            errors = err_buffer.getvalue().strip()
            if logs:
                with st.expander("Processing log", expanded=True):
                    st.code(logs)
            if errors:
                with st.expander("Warnings", expanded=True):
                    st.warning(errors)

            st.download_button(
                label="Download allotment result workbook",
                data=result_buffer.getvalue(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.info("The workbook contains the Allotment, Seats, and Rounds worksheets.")
        except Exception as exc:
            st.error(f"Allotment failed: {exc}")
            st.exception(exc)


if __name__ == "__main__":
    main()
