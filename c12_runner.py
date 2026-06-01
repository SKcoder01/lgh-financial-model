"""
c12_runner.py — flat orchestrator for the C12 model.

Runs the modules in dependency order and threads their results between them
IN MEMORY (no Excel round-trip between modules). The workbook is opened once at
the start; in produce mode it would be written once at the end. In calibrate
mode (the default here) nothing is written.

    departments (leaves)            ->        c12_oc (consolidation)
    c12_hm / c12_d&s / c12_bi&F               consumes the department result dicts
    each returns a results dict               directly; never re-reads their sheets

The composition is deliberately SHALLOW: the runner calls each pure function and
hands the data forward, rather than letting one module call another internally.
That keeps every module independently testable (the basis of the zero-diff checks).
"""
import c12_dept as cd
import c12_oc

DEPTS = [('c12_hm',  '.c12_hm_hr'),
         ('c12_d&s', '.c12_d&s_hr'),
         ('c12_bi&F', '.c12_bi&f_hr')]


def run(excel_file=cd.EXCEL_FILE, verbose=True):
    """Compute the whole C12 model. Returns {sheet: results_dict}.
    Reports each module's calibration (Python vs the sheet's Excel formulas)."""
    wb, wbd, ci, pf, pt = cd.load(excel_file)          # ONE load for the whole run
    all_results = {}
    ok = True

    # --- departments (leaves): each computed independently ---
    dept_results = {}
    for sheet, hr in DEPTS:
        res, pcs, rev_lgh, rev_c12, dupes = cd.compute_dept(wbd, ci, sheet, hr, pf, pt)
        dept_results[sheet] = (res, rev_lgh, rev_c12)   # what c12_oc will consume
        all_results[sheet] = res
        if verbose:
            diffs = cd.diff_o_rows(wbd[sheet], res, pcs)
            ok = ok and not diffs and not dupes
            tag = "OK zero diffs" if not diffs else f"FAIL {len(diffs)} diffs"
            if dupes:
                tag += f"  DUPLICATE KEYS {dupes}"
            print(f"[{sheet:9}] {tag}")
            for f, i, p, e, py, d in diffs[:30]:
                print(f"    {f:30} {i:8} {p:9} excel={e:14,.2f} py={py:14,.2f} d={d:14,.2f}")

    # --- consolidation: department results passed IN MEMORY (not re-read from Excel) ---
    oc_res, oc_pcs = c12_oc.compute_oc(wbd, ci, pf, pt, dept_results=dept_results)
    all_results['c12_oc'] = oc_res
    if verbose:
        diffs = cd.diff_o_rows(wbd['c12_oc'], oc_res, oc_pcs)
        ok = ok and not diffs
        tag = "OK zero diffs" if not diffs else f"FAIL {len(diffs)} diffs"
        print(f"[{'c12_oc':9}] {tag}  (departments in-memory)")
        for f, i, p, e, py, d in diffs:
            print(f"    {f:34} {p:9} excel={e:14,.2f} py={py:14,.2f} d={d:14,.2f}")

    if verbose:
        print("ALL MODULES OK" if ok else "DIFFERENCES FOUND — see above")
    return all_results


if __name__ == '__main__':
    import sys
    run(sys.argv[1] if len(sys.argv) > 1 else cd.EXCEL_FILE)
