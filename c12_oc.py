"""
c12_oc.py — C12 consolidated P&L.

Unlike the departments, c12_oc is not a thin wrapper — it consolidates the three
departments and layers on costs, profit, the SK distributions, a cumulative LOC
balance, and the accounts-receivable / revenue rollups.

It is reproduced by reading LEAF sources and recomputing the chain (rather than
reading c12_oc's own wiring rows), so it stays independent of Excel's formulas:
  - department outputs   : total_c12_*_revenue, total_lgh_*_revenue, lgh_ar_cumm,
                           hr_reimbursable, hr_not_reimbursable, cumm_accrued_participation
  - c12_sk (index 'sk')  : bank_transfers, lease_payments, guarantee_fees, cuda_buyback
  - .c12_oc_hr           : base_salary (index 'sk'), and the four hr_* summaries
  - c12_oc own i-rows    : nonrecur_exp, ho_exp (office_costs/oop_notsk/accounts_payable),
                           inflation_multiplier, compensation_direct, roots_donation, loc_writeoff

Calc is PURE; calibrate() compares vs Excel and WRITES NOTHING.

Note: the compensation row's actual-vs-model IF resolves to the model side in the
projection periods (no actual present), so compensation = SUM of the four comp lines.

The "OOP paid via sk" value originates on c12_sk (ho_exp / index 'oop_sk') and is read
from there by name; c12_oc r31 is merely a link to it, so it carries no label itself.
"""

import openpyxl
import c12_dept as cd

SHEET, HR = 'c12_oc', '.c12_oc_hr'
DEPTS = ['c12_hm', 'c12_d&s', 'c12_bi&F']
NONRECUR = ['infrastructure', 'professional_services', 'hiring_costs',
            'signing_bonus', 'restructuring', 'housing']

def _read(wbd, sheet, pf, pt):
    ws = wbd[sheet]
    pcs = cd.get_period_cols(ws, pf, pt)
    cols = {'opening': cd.get_opening_col(ws, pf)}; cols.update(pcs)
    return cd.read_sheet(ws, cols), pcs

def compute_oc(wbd, ci, pf, pt, dept_results=None):
    """Reproduce the C12 consolidated P&L. PURE.

    dept_results, when supplied by the runner, is {sheet: (results_dict, rev_lgh, rev_c12)}
    carrying the department modules' IN-MEMORY outputs, so c12_oc never re-reads the
    department sheets from Excel on that path. When None (isolated calibration) the
    department values are read from Excel instead. Returns (results, period_cols).
    """
    ws = wbd[SHEET]
    pcs = cd.get_period_cols(ws, pf, pt); cps = list(pcs)
    opening = cd.get_opening_col(ws, pf)

    # leaf sources that always come from Excel: oc i-rows, oc HR summaries, c12_sk inputs
    oc, _ = _read(wbd, SHEET, pf, pt)
    hr, _ = _read(wbd, HR, pf, pt)
    sk, _ = _read(wbd, 'c12_sk', pf, pt)

    # department values: in-memory if the runner passed them, else read from Excel
    if dept_results is None:
        dept_results = {}
        for s in DEPTS:
            data, _ = _read(wbd, s, pf, pt)
            rl, rc = cd.discover_revenue_fields(wbd[s])
            dept_results[s] = (data, rl, rc)

    def D(s, f, p):  return cd.get(dept_results[s][0], f, '', p)
    def RC(s, p):    return cd.get(dept_results[s][0], dept_results[s][2], '', p)   # total_c12_*_revenue
    def RL(s, p):    return cd.get(dept_results[s][0], dept_results[s][1], '', p)   # total_lgh_*_revenue
    def I(f, i, p):  return cd.get(oc, f, i, p)
    def H(f, p, i=''): return cd.get(hr, f, i, p)
    def SK(f, p, i='sk'): return cd.get(sk, f, i, p)

    res = {}
    def store(f, p, v): res.setdefault((f, ''), {})[p] = v
    prior_loc = ws.cell(62, opening).value or 0.0   # opening LOC balance for roll-forward

    for p in cps:
        # participation revenue
        part_rev = sum(RC(s, p) for s in DEPTS)
        store('total_participation_revenue_actual', p, part_rev)

        # compensation (model side: sum of the four comp lines)
        comp = (-(D('c12_hm','hr_reimbursable',p)   + D('c12_hm','hr_not_reimbursable',p))
                -(D('c12_bi&F','hr_reimbursable',p) + D('c12_bi&F','hr_not_reimbursable',p))
                -(D('c12_d&s','hr_reimbursable',p)  + D('c12_d&s','hr_not_reimbursable',p))
                -(H('hr_grand_total',p) - H('hr_only_c12',p)))
        store('compensation_total', p, comp)

        # expenses paid by lgh
        nonrecur = sum(I('nonrecur_exp', idx, p) for idx in NONRECUR)
        office, oop_notsk, ap = I('ho_exp','office_costs',p), I('ho_exp','oop_notsk',p), I('ho_exp','accounts_payable',p)
        oop_sk = SK('ho_exp', p, 'oop_sk')                              # originates on c12_sk
        ho_block = office + oop_sk + oop_notsk + (-nonrecur) + ap        # r30:r34 (r33 = -nonrecur)
        inflation = (I('inflation_multiplier','',p) - 1) * (comp + ho_block)   # 0 when multiplier = 1
        exp_lgh = ho_block + inflation + comp + I('compensation_direct','',p)
        store('total_c12_exp_lgh_actual', p, exp_lgh)

        # revenue, costs, profit
        rev_from_lgh = (-exp_lgh) + part_rev
        exp_not_lgh  = (-H('hr_only_c12',p)) + I('roots_donation','',p)
        profit = rev_from_lgh + (exp_not_lgh + exp_lgh)
        store('total_c12_profit_actual', p, profit)

        # distributions (c12_sk + oc-internal, all index 'sk')
        dist = (SK('bank_transfers',p) + (-oop_sk) + SK('lease_payments',p)
                + H('base_salary',p,'sk') + (-SK('guarantee_fees',p)) + (-SK('cuda_buyback',p)))
        store('distribution_required', p, dist)

        # LOC balance (cumulative, capped at 0)
        loc_bal = min(profit + dist + prior_loc + I('loc_writeoff','',p), 0)
        store('loc_balance', p, loc_bal)
        prior_loc = loc_bal

        # AR / revenue rollups
        store('total_lgh_acc_rec', p, -D('c12_hm','lgh_ar_cumm',p) + D('c12_bi&F','lgh_ar_cumm',p) - D('c12_d&s','lgh_ar_cumm',p))
        store('total_c12_acc_rec', p, sum(D(s,'cumm_accrued_participation',p) for s in DEPTS))
        store('total_lgh_revenue', p, sum(RL(s,p) for s in DEPTS))

    return res, pcs

def calibrate(excel_file=cd.EXCEL_FILE, tol=0.5, verbose=True, dept_results=None):
    """Isolated check: compute vs c12_oc's own Excel formulas. Pass dept_results to
    validate the in-memory consolidation path (departments handed over as Python dicts)."""
    wb, wbd, ci, pf, pt = cd.load(excel_file)
    res, pcs = compute_oc(wbd, ci, pf, pt, dept_results)
    diffs = cd.diff_o_rows(wbd[SHEET], res, pcs, tol)
    if verbose:
        src = "departments in-memory" if dept_results is not None else "departments from Excel"
        print(f"[{SHEET}] periods {list(pcs)}  ({src})")
        if diffs:
            print(f"  FAIL {len(diffs)} differences vs Excel:")
            for f, i, p, e, py, d in diffs:
                print(f"    {f:34} {p:9} excel={e:14,.2f} py={py:14,.2f} d={d:14,.2f}")
        else:
            print("  OK zero differences vs Excel")
    return res, diffs

if __name__ == '__main__':
    calibrate()
