# C12 Financial Model — File Guide & Finance Update Instructions

## How the model works (in one paragraph)

The C12 model lives in an Excel workbook (`c12_io.xlsx`). Finance works **only in
Excel** — entering actuals and assumptions. A Python program (the "runner") reads
those inputs, calculates the model, and writes the results back into the workbook.
So Excel is the **interface** (what you type and what you read) and Python is the
**calculation engine**. You never run Python; you update the workbook and hand it
back, and the operator runs the model.

---

## The files

### Workbooks
- **`c12_io.xlsx`** — the working C12 model. **This is the file finance updates.**
  It is colour-coded (see below) so it is obvious which cells to change.
- **`financial_model_excel_claude.xlsx`** — the full master model (the LGH model,
  the property sheets, and the C12 model together). This is the reference copy, not
  the finance working file. Do not edit this one.

### Python engine (run by the operator, not finance)
- **`c12_runner.py`** — runs the whole C12 model in one go.
- **`c12_dept.py`** — the shared calculation engine used by every department.
- **`c12_hm.py`, `c12_ds.py`, `c12_bif.py`** — the three department configs
  (Hotel Management, Design & Structuring, Business Infrastructure & Finance).
- **`c12_oc.py`** — the consolidation (combines departments into the C12 P&L).

### Helper scripts (operator)
- **`produce.py`** — writes the model's calculated results into the workbook.
- **`backtest.py`** — checks the model's output against recorded actuals for any
  period that has both (a self-consistency / accuracy check).
- **`save_clean.py`** — ensures a workbook still opens cleanly in Excel after Python
  has written to it.
- **`check_model.py`** — diagnostic check for the full master model.
- **`colour_finance.py`** — re-applies the finance colour scheme to the workbook.

### Documentation
- **`context_document.md`** — the full technical reference for the model.
- **`C12_MODEL_GUIDE.md`** — this file.

---

## What finance updates in `c12_io.xlsx`

### The only thing you need to remember

| Colour | Meaning | Action |
|--------|---------|--------|
| 🟢 **Green** | Your inputs and actuals | **Update these** |
| ⬜ **Grey** | Model outputs | **Leave alone** — the model fills these in |
| 🟡 Yellow | Labels (row names, field names) | Don't change |
| 🟤 Brown | Dates / period headers | Don't change |
| ⬜ White | Internal calculations | Don't change |

**Rule of thumb: only ever type into green cells.** Grey cells will be blank or
out of date until the model is run — that is expected; the operator's run fills them.

### Where the green cells are

- **Actuals (left-hand columns):** the recorded actual figures for periods that have
  already happened, one column per quarter.
- **Model inputs (right-hand columns):** the assumptions/inputs for the projected
  quarters.
- **Input pages — `c12_inputs` and `.c12_inputs.hr`:** the model-wide assumptions
  (rates, percentages, salaries, and the control dates below).
- **`.c12_alloc`:** the staff time-allocation matrix. Updated occasionally, not every
  period.

### Pages you should know about

- **Visible `c12_…` pages** — the model itself; update the green cells.
- **`c12_bs`** — balance-sheet view. It is **display only** (all white) and pulls its
  figures from the other pages. Do not edit it.
- **Hidden `.c12_…` pages** (HR/salary detail) — sensitive. Only update green cells
  here if you are responsible for that data.
- **`registry` and `c12_diags`** — system/diagnostic pages. Do not touch.

### The three control dates (on `c12_inputs`)

These tell the model which periods are history and which to calculate. They are green
(editable) but are normally set by the modeller, not by finance — check before changing:

- **`actuals_thru`** — the last period that has real actuals (currently 31 Mar 2026).
- **`projections_from`** — the first period the model calculates (currently 30 Jun 2026).
- **`projections_to`** — the last period the model calculates (currently 31 Dec 2026).

---

## The update cycle

1. Finance updates the **green** cells (actuals and assumptions).
2. Save the workbook and return it.
3. The operator runs `c12_runner` — this recalculates the model and fills the **grey**
   output cells with the results.
4. The updated workbook comes back with the calculated results in place.

---

## A few cautions

- **Don't add or delete columns or rows** without checking with the modeller. The model
  finds data by its label (not its position), so inserting a normal row is usually safe,
  but changing the **period columns** needs care.
- **When you open the file it may recalculate for a moment** — that is normal.
- **If a grey cell looks wrong or empty,** that's because the model hasn't been run since
  the last edit — it isn't an error. Running the model refreshes it.
