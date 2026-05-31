"""
c12_hm.py — Hotel Management department model
Reads inputs from Excel, calculates, writes outputs back to Excel.

Inputs:
  c12_hm      (i rows, projection cols defined by projections_from/projections_to dates)
  c12_hm      (i rows, col I = Q4 2025 actual = opening balance)
  .c12_hm_hr  (o rows: hr_grand_total, hr_reimbursable, hr_not_reimbursable, hr_only_c12)
  c12_inputs  (scalar assumptions + date controls)

Outputs written to:
  c12_hm      (o rows, projection period cols)

Python does NOT touch .c12 sheets — HR summary rows are read as inputs only.
Excel handles all .c12 calculations internally.

Period range is dynamic — driven by projections_from and projections_to in c12_inputs.
Opening balance: col I (Q4 2025 actual).
"""

import openpyxl
from datetime import date, datetime

# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = 'financial_model_excel_claude.xlsx'

# ============================================================
# DYNAMIC COLUMN DETECTION — uses row 3 markers
#
# Row 3 convention:
#   "actual" appears in the first column of the actuals section
#   "i/o"    appears in the first column of the model i/o section
#
# This means adding actual columns or comment columns never
# breaks Python — it just reads the markers to find boundaries.
# ============================================================

def find_section_boundaries(ws):
    """
    Scan row 3 for 'actual' and 'i/o' markers.
    Returns: (actual_start_col, io_start_col)
    """
    actual_start = None
    io_start = None
    for col in range(1, ws.max_column + 1):
        v = ws.cell(3, col).value
        if isinstance(v, str):
            if v.strip().lower() == 'actual' and actual_start is None:
                actual_start = col
            elif v.strip().lower() == 'i/o' and io_start is None:
                io_start = col
    return actual_start, io_start

def get_opening_col(ws, projections_from):
    """
    Find the opening balance column — the last date in the actuals
    section (cols before i/o marker) that is <= projections_from.
    """
    _, io_start = find_section_boundaries(ws)
    if not io_start:
        return None
    # Search actuals cols (before io_start) for latest date <= projections_from
    best_col = None
    best_date = None
    for col in range(1, io_start):
        v = ws.cell(2, col).value
        if not isinstance(v, (date, datetime)):
            continue
        d = v.date() if isinstance(v, datetime) else v
        if d < projections_from:
            if best_date is None or d > best_date:
                best_date = d
                best_col = col
    return best_col

def get_period_cols(ws, from_date, to_date):
    """
    Find model i/o columns between from_date and to_date.
    Only looks in cols from the i/o marker onwards.
    Returns: {'Q2_2026': 17, 'Q3_2026': 18, ...} sorted by col.
    """
    _, io_start = find_section_boundaries(ws)
    if not io_start:
        return {}
    cols = {}
    for col in range(io_start, ws.max_column + 1):
        v = ws.cell(2, col).value
        if not isinstance(v, (date, datetime)):
            continue
        d = v.date() if isinstance(v, datetime) else v
        if from_date <= d <= to_date:
            q = (d.month - 1) // 3 + 1
            label = f"Q{q}_{d.year}"
            cols[label] = col
    return dict(sorted(cols.items(), key=lambda x: x[1]))

def get_actuals_col(ws, period_label):
    """
    Find the actual column for a period label.
    Only searches in actuals section (before i/o marker).
    """
    _, io_start = find_section_boundaries(ws)
    limit = io_start if io_start else ws.max_column + 1
    for col in range(1, limit):
        v = ws.cell(2, col).value
        if not isinstance(v, (date, datetime)):
            continue
        d = v.date() if isinstance(v, datetime) else v
        q = (d.month - 1) // 3 + 1
        label = f"Q{q}_{d.year}"
        if label == period_label:
            return col
    return None

# ============================================================
# HELPER: READ SHEET INTO DICT
# Reads all rows by field_name+index, for specified columns
# ============================================================

def read_sheet(ws, cols, io_col=2, field_col=3, index_col=4):
    """
    Read specified columns from a sheet.
    Returns: {(field_name, index): {col_key: value}}
    cols: dict of {label: col_number}
    """
    data = {}
    for r in range(1, ws.max_row + 1):
        io = ws.cell(r, io_col).value
        if not isinstance(io, str) or io.strip().lower() not in ('i', 'o', 'c'):
            continue
        field = ws.cell(r, field_col).value
        index = ws.cell(r, index_col).value or ''
        if not field:
            continue
        key = (field, index)
        data[key] = {'io': io.strip().lower()}
        for label, col in cols.items():
            v = ws.cell(r, col).value
            data[key][label] = v if isinstance(v, (int, float)) else 0.0
    return data

def get(data, field, index='', period='Q1_2026', default=0.0):
    """Lookup value from data dict."""
    key = (field, index)
    if key not in data:
        return default
    return data[key].get(period, default)

def get_opening(data, field, index='', default=0.0):
    """Get opening balance (Q4 2025 actual)."""
    return get(data, field, index, 'opening', default)

def prior(results, field, period, calc_periods, default=0.0):
    """Get prior period result."""
    idx = calc_periods.index(period)
    if idx == 0:
        return get_opening(results, field, '', default)
    return results.get((field,''), {}).get(calc_periods[idx-1], default)

# ============================================================
# READ INPUTS
# ============================================================

def read_c12_inputs(wb):
    """Read scalar assumptions and date controls from c12_inputs."""
    ws = wb['c12_inputs']
    inputs = {}
    for r in range(1, ws.max_row + 1):
        io = ws.cell(r, 2).value
        if io != 'i': continue
        field = ws.cell(r, 3).value
        val = ws.cell(r, 5).value
        if not field: continue
        if isinstance(val, (int, float)):
            inputs[field] = val
        elif isinstance(val, datetime):
            inputs[field] = val.date()
        elif isinstance(val, date):
            inputs[field] = val
    return inputs

def read_hm_inputs(wb_data, period_cols, opening_col):
    """Read c12_hm i rows — opening balance + projection period values."""
    ws = wb_data['c12_hm']
    cols = {'opening': opening_col}
    cols.update(period_cols)
    return read_sheet(ws, cols)

def read_hr_outputs(wb, period_cols, opening_col, wb_data=None):
    """
    Read the four HR summary o rows from .c12_hm_hr.
    These are calculated by Excel — Python reads them as inputs.
    Uses data_only workbook to get calculated formula values.
    """
    ws = (wb_data or wb)['.c12_hm_hr']
    hr_fields = {'hr_grand_total', 'hr_reimbursable', 'hr_not_reimbursable', 'hr_only_c12'}
    cols = {'opening': opening_col}
    cols.update(period_cols)
    data = read_sheet(ws, cols)
    return {k: v for k, v in data.items() if k[0] in hr_fields}

# ============================================================
# CHECK ACTUALS
# ============================================================

def get_value(inputs, hr_data, field, index, period, actuals_through=None):
    """
    Return actual if period <= actuals_through, else projection input.
    For o rows: Python calculates, but actual overrides if exists.
    """
    key = (field, index)
    proj = (inputs.get(key) or hr_data.get(key) or {}).get(period, 0.0)

    if actuals_through and period == 'Q1_2026':
        # Check col J for Q1 2026 actual
        # actuals_through date determines if actual exists
        pass  # handled per-row in main calc

    return proj

# ============================================================
# MAIN CALCULATION: c12_hm
# ============================================================

def calc_hm(inputs, hr_data, c12_inputs, calc_periods, actuals_through=None):
    """
    Calculate all c and o rows for c12_hm.
    calc_periods: list of period labels in order e.g. ['Q1_2026', 'Q2_2026', ...]
    Returns: {(field, index): {period: value}}
    """
    chargeback_pct    = c12_inputs.get('chargeback_pct_actual', 0.05)
    participation_pct = c12_inputs.get('participation_pct_actual', 0.35)
    chargeback_mult   = 1 / (1 - chargeback_pct)

    results = {}

    def store(field, index, period, val):
        key = (field, index)
        if key not in results:
            results[key] = {}
        results[key][period] = val

    def inp(field, index='', period='Q1_2026'):
        return get(inputs, field, index, period)

    def hr(field, period='Q1_2026'):
        return get(hr_data, field, '', period)

    def res(field, period, default=0.0):
        """Get result from current calculation."""
        return results.get((field,''), {}).get(period, default)

    def prior_res(field, period):
        """Get prior period result, using opening balance for first period."""
        idx = calc_periods.index(period)
        if idx == 0:
            return get(inputs, field, '', 'opening', 0.0)
        return results.get((field,''), {}).get(calc_periods[idx-1], 0.0)

    properties = ['bermuda_pb', 'austin', 'turks', 'fritholme', 'elbow', 'lsat', 'l4']
    fee_fields = ['mgmt_fee', 'mgmt_fee_actual', 'incentive_fee', 'incentive_fee_actual']

    for period in calc_periods:

        # ---- FEE REVENUE ----
        total_fees_earned = sum(
            inp(f, prop, period)
            for f in fee_fields
            for prop in properties
        )
        deferred_fee      = inp('deferred_fee', '', period)
        paid_fee_accruals = inp('paid_fee_accruals', '', period)
        cumm_accrued_fee  = deferred_fee + prior_res('cumm_accrued_fee', period) + paid_fee_accruals
        paid_current_fees     = total_fees_earned + deferred_fee
        total_fees_paid       = paid_fee_accruals + paid_current_fees

        store('total_fees_earned',  '', period, total_fees_earned)
        store('cumm_accrued_fee',   '', period, cumm_accrued_fee)
        store('paid_current_fees',  '', period, paid_current_fees)
        store('total_fees_paid',    '', period, total_fees_paid)

        # ---- SERVICE REVENUE ----
        total_service_revenue = sum(inp('service_fee', prop, period) for prop in properties)
        deferred_service_fee      = inp('deferred_service_fee', '', period)
        paid_service_accruals     = inp('paid_service_accruals', '', period)
        cumm_accrued_service_fee  = deferred_service_fee + prior_res('cumm_accrued_service_fee', period) + paid_service_accruals
        paid_current_service_fees    = total_service_revenue + deferred_service_fee
        total_service_revenue_paid   = paid_service_accruals + paid_current_service_fees

        store('total_service_revenue',       '', period, total_service_revenue)
        store('cumm_accrued_service_fee',    '', period, cumm_accrued_service_fee)
        store('paid_current_service_fees',   '', period, paid_current_service_fees)
        store('total_service_revenue_paid',  '', period, total_service_revenue_paid)

        # ---- CHARGEBACKS ----
        hr_reimbursable        = hr('hr_reimbursable', period)
        reimbursable_by_hotels = hr_reimbursable
        total_chargeback       = reimbursable_by_hotels * chargeback_mult

        store('chargeback_multiplier',   '', period, chargeback_mult)
        store('reimbursable_by_hotels',  '', period, reimbursable_by_hotels)
        store('total_chargeback',        '', period, total_chargeback)

        # ---- LGH REVENUE TOTALS ----
        total_lgh_hm_revenue = total_service_revenue_paid + total_fees_paid + total_chargeback
        lgh_ar_cumm          = cumm_accrued_fee + cumm_accrued_service_fee

        store('total_lgh_hm_revenue', '', period, total_lgh_hm_revenue)
        store('lgh_ar_cumm',          '', period, lgh_ar_cumm)

        # ---- C12 PARTICIPATION ----
        pre_mod_extra_pct  = inp('pre_mod_extra_pct',  '', period)
        catchup_extra_pct  = inp('catchup_extra_pct',  '', period)

        c12_participation_amt = (
            (participation_pct + pre_mod_extra_pct) * (paid_current_fees + paid_current_service_fees) +
            (participation_pct + catchup_extra_pct) * (paid_fee_accruals + paid_service_accruals)
        )
        chargeback_fees       = chargeback_pct * total_chargeback
        total_c12_hm_revenue  = chargeback_fees + c12_participation_amt

        store('c12_participation_pct',  '', period, participation_pct)
        store('chargeback_pct',         '', period, chargeback_pct)
        store('c12_participation_amt',  '', period, c12_participation_amt)
        store('chargeback_fees',        '', period, chargeback_fees)
        store('total_c12_hm_revenue',   '', period, total_c12_hm_revenue)

        increase = -(participation_pct + pre_mod_extra_pct) * (deferred_fee + deferred_service_fee)
        decrease = -(participation_pct + catchup_extra_pct) * (paid_fee_accruals + paid_service_accruals)
        store('increase_accrued_participation', '', period, increase)
        store('decrease_accrued_participation', '', period, decrease)

        # ---- HR (read from .c12_hm_hr outputs) ----
        store('hr_grand_total',      '', period, hr('hr_grand_total',      period))
        store('hr_reimbursable',     '', period, hr('hr_reimbursable',     period))
        store('hr_not_reimbursable', '', period, hr('hr_not_reimbursable', period))
        store('hr_only_c12',         '', period, hr('hr_only_c12',         period))

    return results

# ============================================================
# WRITE OUTPUTS BACK TO EXCEL
# ============================================================

def write_outputs(ws, results, period_cols, actuals_through=None):
    """Write o row results back to projection columns in c12_hm."""
    written = 0
    for r in range(1, ws.max_row + 1):
        io = ws.cell(r, 2).value
        field = ws.cell(r, 3).value
        index = ws.cell(r, 4).value or ''
        if io != 'o' or not field:
            continue
        key = (field, index)
        if key not in results:
            continue
        for period, col in period_cols.items():
            # Skip if actual exists for this period (check actuals cols E-K)
            actual_col = get_actuals_col(ws, period)
            if actual_col and actuals_through:
                actual_val = ws.cell(r, actual_col).value
                if isinstance(actual_val, (int, float)):
                    continue  # actual exists — don't overwrite
            val = results[key].get(period)
            if val is not None:
                ws.cell(r, col).value = round(val, 2)
                written += 1
    return written

# ============================================================
# RUN
# ============================================================

def run(excel_file=EXCEL_FILE):
    print(f"Reading: {excel_file}")
    wb      = openpyxl.load_workbook(excel_file)
    wb_data = openpyxl.load_workbook(excel_file, data_only=True)

    c12_inputs = read_c12_inputs(wb)
    actuals_through  = c12_inputs.get('actuals_thru') or c12_inputs.get('actuals_through')
    projections_from = c12_inputs.get('projections_from')
    projections_to   = c12_inputs.get('projections_to')

    print(f"  actuals_through:  {actuals_through}")
    print(f"  projections_from: {projections_from}")
    print(f"  projections_to:   {projections_to}")

    # Detect column boundaries from row 3 markers
    ws_hm_data = wb_data['c12_hm']
    opening_col  = get_opening_col(ws_hm_data, projections_from)
    period_cols  = get_period_cols(ws_hm_data, projections_from, projections_to)
    calc_periods = list(period_cols.keys())

    print(f"  Opening balance col: {opening_col}")
    print(f"  Periods: {calc_periods}")

    hm_inputs = read_hm_inputs(wb_data, period_cols, opening_col)
    hr_data   = read_hr_outputs(wb, period_cols, opening_col, wb_data)

    print(f"  HM input rows: {len(hm_inputs)}")
    print(f"  HR summary rows: {len(hr_data)}")

    results = calc_hm(hm_inputs, hr_data, c12_inputs, calc_periods, actuals_through)
    print(f"  Calculated: {len(results)} fields")

    ws_hm_write = wb['c12_hm']
    written = write_outputs(ws_hm_write, results, period_cols, actuals_through)
    print(f"  Written to c12_hm: {written} cells")

    wb.save(excel_file)
    print(f"  Saved: {excel_file}")

    return results

if __name__ == '__main__':
    run()
