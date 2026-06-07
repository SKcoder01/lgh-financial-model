"""
backtest.py — model-vs-actual backtest for the C12 model.

Idea: set projections_from to a date that ALSO has recorded actuals, run the
model, and compare each computed o-value against the actual column for the SAME
date. A correct model that's been fed consistent inputs reproduces the actuals
exactly. Any gap is a model-vs-actual variance to investigate (or, before real
actuals are loaded, model drift vs whatever was pasted into the actuals block).

Generalized over ALL overlapping periods, not just one quarter: for each model
period the model computes, it looks for an actuals-block column with a matching
date and compares. Periods with no matching actual column are skipped.

Usage:
    import backtest
    backtest.run('c12_io.xlsx', projections_from='2026-03-31')   # override window
    backtest.run('c12_io.xlsx')                                  # use file's own date

Reads a cache-fresh (Excel-saved) workbook. Writes nothing.
"""
import datetime
import c12_dept as cd
import c12_oc

DEPTS = [('c12_hm', '.c12_hm_hr'), ('c12_d&s', '.c12_d&s_hr'), ('c12_bi&F', '.c12_bi&f_hr')]

def _as_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date):     return v
    if isinstance(v, str):               return datetime.date.fromisoformat(v)
    return v

def _actual_cols_by_date(ws):
    """Map {date: column} for the ACTUALS block (columns left of io_start)."""
    _, io_start = cd.find_section_boundaries(ws)
    limit = io_start if io_start else ws.max_column + 1
    out = {}
    for c in range(1, limit):
        v = ws.cell(2, c).value
        d = _as_date(v) if isinstance(v, (datetime.date, datetime.datetime)) else None
        if d is not None and d not in out:   # first (leftmost) column for that date
            out[d] = c
    return out

def _period_label_to_date(label, period_cols, ws):
    """Resolve a period label (e.g. 'Q1_2026') back to its date via its model column."""
    c = period_cols[label]
    return _as_date(ws.cell(2, c).value)

def run(excel_file, projections_from=None, tol=0.5, verbose=True):
    wb, wbd, ci, pf, pt = cd.load(excel_file)
    if projections_from is not None:
        pf = _as_date(projections_from)
        ci['projections_from'] = pf

    # compute the whole model under this window
    dept_results = {}; pcs_map = {}
    for sheet, hr in DEPTS:
        res, pcs, rl, rc, dup = cd.compute_dept(wbd, ci, sheet, hr, pf, pt)
        dept_results[sheet] = (res, rl, rc); pcs_map[sheet] = (res, pcs)
    oc_res, oc_pcs = c12_oc.compute_oc(wbd, ci, pf, pt, dept_results=dept_results)
    pcs_map['c12_oc'] = (oc_res, oc_pcs)

    if verbose:
        print(f"Backtest: projections_from={pf}  projections_to={pt}")
        print("Comparing each model o-value against the actual column of the same date.\n")

    report = {}; all_ok = True
    for sheet, (res, pcs) in pcs_map.items():
        ws = wbd[sheet]
        act_by_date = _actual_cols_by_date(ws)
        diffs = []; compared = 0
        for r in range(1, ws.max_row + 1):
            if ws.cell(r, 2).value != 'o':
                continue
            f = ws.cell(r, 3).value; i = ws.cell(r, 4).value or ''
            if not f or (f, i) not in res:
                continue
            for label in pcs:
                pdate = _period_label_to_date(label, pcs, ws)
                acol = act_by_date.get(pdate)
                if acol is None:
                    continue                       # no actual column for this date
                actual = ws.cell(r, acol).value
                if not isinstance(actual, (int, float)):
                    continue                       # blank actual -> nothing to compare
                model = res[(f, i)].get(label)
                if model is None:
                    continue
                compared += 1
                if abs(model - actual) > tol:
                    diffs.append((f, i, label, pdate, model, actual, model - actual))
        report[sheet] = diffs
        all_ok = all_ok and not diffs
        if verbose:
            tag = "OK" if not diffs else f"{len(diffs)} mismatch(es)"
            print(f"  [{sheet:9}] {compared:3} model/actual pairs compared  -> {tag}")
            for f, i, label, pdate, m, a, d in diffs:
                print(f"      {f:32} {i:6} {pdate} model={m:14,.1f} actual={a:14,.1f} diff={d:,.1f}")
    if verbose:
        print("\nMODEL REPRODUCES ACTUALS on every overlapping period."
              if all_ok else "\nVARIANCES found above (model vs actual).")
    return report
