# LGH/C12 Financial Model — Claude Context Document
## Version: 2.0 — June 2026
## Purpose: Paste at start of new Claude conversation to restore full context

> **v2.0 update note:** The entire C12 Python model is now built, validated three
> ways, and proven extensible to new periods with no code change. Several hard-won
> openpyxl facts and a reusable toolset (`save_clean`, `produce`, `backtest`,
> `colour_finance`, `run_model`) were added. Exact procedures (colour coding,
> corruption fix, running the model) now live in committed scripts — this document
> points to them rather than re-describing them, so a fresh conversation reproduces
> them exactly by running the code, not by re-deriving it. See sections 7, 8, 10, 11,
> 17, 18.

---

## 1. PROJECT OVERVIEW

Converting the LGH/C12 Capital Excel financial model to Python.
GitHub repo: https://github.com/SKcoder01/lgh-financial-model.git

**Two computers:**
- Old laptop: `C:\Users\StephenWrkbook-C12\Desktop\excel_claude_financial model\`
- New laptop: `C:\Users\steph\Desktop\lgh-financial-model\`

**Tech stack:** Python + openpyxl. Never use xlwings (Snapdragon incompatibility on old laptop).

---

## 2. ENTITIES

- **LGH** — investment vehicle, consolidates property performance
- **C12** — operating company, fee revenues, staff costs, departmental P&L
- **5 properties:** Pink Beach (Bermuda), Elbow Beach (Bermuda), Austin Lady Bird Lake, Turks & Caicos, Fritholme, San Antonio (LSAT — smallest)
- **3 C12 departments:** Hotel Management (HM), Design & Structuring (D&S), Business Infrastructure & Finance (Bi&F)

---

## 3. PHASING

**Phase 1a (current):**
- Python handles lgh and c12 calculations
- cs2 property sheets stay in Excel
- Registry manages data transfer
- Data entered manually by finance team per metacode
- Code changes done by SK with Claude

**Phase 1b:**
- cs2s built in Excel, converted to Python by Claude and tested
- New hire works with Claude on c12/lgh modifications

**Phase 2:**
- M3/Opera data integrated as feeds
- Salesforce integrated
- Manual entry eliminated

---

## 4. ARCHITECTURE

### Folder Model Pattern
Single Excel file + Python scripts in one folder.

### Sheet Naming Conventions
- `lgh_*` — LGH model sheets
- `c12_*` — C12 model sheets  
- `pink_cs2`, `elbow_cs2`, `aus_cs2`, `turks_cs2`, `frith_cs2`, `sat_cs2` — property models
- `.c12_*_hr` — HR/salary data (hidden, locked, sensitive)
- `.c12_alloc` — time allocation matrix
- `.c12_inputs.hr` — HR assumption inputs
- `registry` — interface between cs2s and lgh/c12
- `lgh_diags` — diagnostic comparison sheet

### Cross-Sheet Reference Rules
- Same prefix sheets can reference each other directly (lgh↔lgh, c12↔c12)
- Cross-prefix must go via registry
- `c12_##` can read from its own `.c12_##_hr` but not vice versa
- `.c12_##_hr` can read from `.c12_inputs.hr`

### Registry
- System tool — users never touch it
- Interface between cs2 property sheets and lgh/c12 calculations
- Cells read by registry have **thick bold border** and **bold text**

### Calculation mode: CALIBRATE vs PRODUCE (important)
- **Calibrate (default):** Python computes everything in memory, compares against the
  workbook's existing `o`-values, prints OK/diffs, and **writes nothing**. Use to verify
  the model still reproduces the file after edits.
- **Produce:** Python computes and **writes** the `o`-values into the o-cells — replacing
  their formulas with hard numbers — for the projection window only. This is the run that
  "fills the greys". Pre-projection/seed columns keep their formulas.
- Both run from `run_model.py` (see section 17).

### Inter-module data transfer is IN-MEMORY (not via Excel)
The runner computes the three departments as pure functions returning result dicts, then
threads those dicts straight into `c12_oc` in memory. Excel is touched only at the
boundaries (read inputs at start, write outputs at end). Reason: openpyxl loads→modifies→
saves but does NOT evaluate formulas, so routing values back through Excel between modules
would force a save/reload/recalc each time. Composition is kept SHALLOW (runner orchestrates
pure functions side by side, not nested) so each module still calibrates in isolation.

---

## 5. I/O/C FRAMEWORK

Every row in every c12/lgh sheet has cols B, C, D:
- **Col B** = `i/o/c` marker
- **Col C** = `field_name` (Python-friendly: lowercase, underscores)
- **Col D** = `index` (property name or person code where matrix)

**Markers:**
- `i` — Python reads this value as input from Excel
- `o` — Python calculates and writes result back to Excel
- `c` — internal Python calculation only, never written to Excel

**Rules:**
- Actuals only allowed on `i` and `o` rows, never `c` rows
- Python uses field_name + index as unique key (never row numbers)
- Adding rows in Excel doesn't break Python — it finds rows by name

**Refinements (learned in build):**
- An `i` row may legitimately contain an Excel *formula* (finance convenience, e.g.
  `=745476/4`, or escalation chains like `=+I7*1.03`). Python reads only the cell VALUE
  and must NEVER replicate or capture an i-row's formula. Flag i-rows-with-formulas
  informationally only; never "correct" them.
- A value that crosses a module boundary (computed by one module, consumed by another)
  must be marked `o` — so it is validated by calibration and written — NOT `i` or `c`.
  Example caught & fixed: `cumm_accrued_participation` was mislabelled `i` (it had a SUM
  formula) but is model logic feeding `c12_oc`; flipped to `o` on c12_hm/d&s/bi&F.
- Period-on-period growth lives ENTIRELY in the inputs. The engine is deterministic with
  no built-in escalation: flat inputs → flat flow rows; only roll-forwards
  (loc_balance, cumm, AR cumulatives) move on flat inputs because they accumulate. To make
  a forecast grow you either keep i-cells as escalating formulas or enter pre-grown hard
  numbers. (Relevant for LGH out-years: flat copies will look flat unless growth is in the inputs.)

---

## 6. COLUMN STRUCTURE (all c12 sheets)

### STANDING CONVENTION (applies to every page, lgh too)
- **Row 2 = dates** (the period axis), always.
- **Row 3 = section markers** (`actual` / `i/o`), always.
- Python keys on this: `find_section_boundaries` reads row 3; period selection reads row 2.
  Keep every page uniform so the engine treats them identically.

### Row 3 Markers (critical for Python)
- `"actual"` — appears in first column of actuals section
- `"i/o"` — appears in first column of model i/o section
- Python scans row 3 to find these markers dynamically. If the `i/o` marker is missing,
  column detection returns None and the runner errors in read_sheet — this actually
  happened when the marker was accidentally deleted; the fix is to restore `i/o` in row 3
  of the first model column on every page.

### Two live file layouts (both valid, different horizons)
- **Full master (`financial_model_excel_claude.xlsx`):** 36 sheets. Actuals E + F:K,
  model block from col **L**, plus annual columns 2027–2030 to the right. `actual`@E, `i/o`@L.
- **Reduced production cut (`c12_io.xlsx`):** 15 sheets (c12 pages + registry + c12_diags
  only; lgh/cs2 stripped). Quarterly-only, history chopped, **2025Q4 hardwired as the seed**,
  annual columns deleted (from visible AND hidden HR pages). `actual`@E, `i/o`@**I**. Model
  block I:M = 2025Q4(seed)→2026Q4. The extended variant runs to 2027Q4 (cols N:Q).

### Column Zones (reduced c12_io layout)
- **Col A** — row label
- **Cols B-D** — i/o/c, field_name, index (metadata)
- **Cols E-H** — actuals zone (E=2024YE, then quarters; actuals_thru currently 2026-03-31)
- **Cols I onwards** — model i/o projection columns (i/o zone). I = seed/opening.

### Date Controls (in c12_inputs)
- `actuals_thru` (E11) — last period with actual data (date)
- `projections_from` (E12) — first period Python calculates (date)
- `projections_to` (E13) — last period Python calculates (date)

**Opening balance** = last actuals col with date strictly < projections_from

### What the row-2 DATE is actually used for (and the known weakness)
The date drives three things in `c12_dept.py`: (1) `get_period_cols` selects which model
columns fall in [from_date, to_date]; (2) `get_opening_col` finds the seed column (latest
date < projections_from); (3) the period LABEL `Q{q}_{year}` is derived from the date and
is how a model column is matched to its same-dated actuals column (backtest depends on this).
**Weakness:** extending the horizon currently requires hand-typing the new period dates into
row 2 — a correctness dependency (a mistype silently mislocates periods). DECISION (SK): the
date is generated in Excel by the data-inputter; making the input surface robust is a
separate concern, deliberately parked. The engine just reads what's there.

---

## 7. COLOUR CODING

Two distinct schemes exist. **The exact rules and hex codes live in the committed
scripts — run them, don't re-derive.** This section records intent and which to use.

### Finance-facing scheme — `colour_finance.py` (current, for the io deliverable)
Built around one instruction to finance: **GREEN = update, GREY = leave alone.**
- brown = row 2 (dates) and row 3 (markers)
- yellow = text labels / metadata (col A, i/o, field_name, index)
- green = actuals (hard numbers left of the `i/o` marker) + model inputs (`i`-rows in the
  i/o block) + input-page input values
- grey = outputs (`o`-rows in the i/o block)
- white = everything else (`c`-rows, calcs, empties, formulas in the actuals zone)
- Boundary derived from the row-3 `i/o` marker (auto-adjusts if columns shift).
- `c12_bs` comes out white (untagged display page); `.c12_alloc` comes out green
  (finance maintains it, infrequently) — both confirmed correct by SK.

### Diagnostic scheme — `colour.py` (for SK's own inspection)
Richer: adds orange (IF formulas), light-blue (actuals), etc. Used as a change-detector
when eyeballing structure. Rules apply in priority order (header → IF → text →
actuals-numeric → i/o/c-by-marker), so an `o`-cell containing an IF shows orange, not grey.

### To apply either: run the script, then it saves via `save_clean` (see section 8).
Applying fills is an openpyxl save, so it MUST go through save_clean or the file will throw
Excel's repair prompt. Both scripts already do this.

---

## 8. DIAGNOSTIC ROUTINES

### "Run Excel Diags" means:
1. `check_model.py` — error matrix in lgh_diags cols X-AG all zeros
2. Cross-prefix violation check — no violations
3. Actuals integrity — no `c` rows have values in actual columns
4. Colour coding reapplied

### check_model.py
Checks `lgh_diags` error matrix (cols X-AG = cols 24-33) for non-zero values.
All zeros = model matches baseline ✓

### Updating baseline
When model changes intentionally: load file data_only, copy live cols to baseline cols in
the diags sheet. (c12_io uses `c12_diags`; full master uses `lgh_diags`. The c12_diags page
was extended to 2027 with LIVE/BASELINE/ERROR blocks; baseline for a new period is frozen to
its first computed result.)

**IMPORTANT — diags compares a file to its OWN frozen baseline, not two files to each other.**
To check whether two files agree (e.g. an Excel-formula version vs a Python-produced version),
do a direct cell-by-cell o-value diff in Python, not the diags page.

---

## 8b. HARD-WON openpyxl FACTS (read before any file surgery)

- **openpyxl save strips cached formula values.** After openpyxl writes a workbook,
  `data_only` reads return None for every formula cell until Excel reopens and recalculates.
  Therefore: ALWAYS calibrate/backtest against a file last saved by **Excel** (fresh cache),
  never one Python just saved. A "FAIL" wall where every diff is `excel=0.00` vs a non-zero
  Python value is this false alarm, not real breakage — open in Excel, recalc, re-check.
- **Empty-`<v>` corruption + the fix.** openpyxl writes formula cells as `<f>…</f><v></v>`
  (empty cached-value tag). Excel sometimes rejects these → "Removed Records: Formula" repair
  prompt, silently deleting formulas. **Fix = `save_clean.py`**: after the openpyxl save it
  strips empty `<v></v>`/`<v/>` after formulas, sets `fullCalcOnLoad="1"`, and drops calcChain
  so Excel rebuilds on open. EVERY Python save of a workbook should go through `save_clean`.
- **openpyxl delete_cols / insert_rows do NOT translate formula references** (silent
  mispointing). Excel does translate. So structural column/row edits are safest done in Excel,
  or in openpyxl only after verifying no surviving formula references the deleted/shifted range.
- **Dragging a formula across columns = `openpyxl.formula.translate.Translator`.** It
  replicates Excel's fill handle exactly: shifts relative refs by the column offset, keeps
  `$`-absolute refs fixed, shifts cross-sheet refs too. Used to extend o/c formulas to new
  periods. No INDIRECT/OFFSET exist in this model, so Translator is safe here.
- **Comments ARE preserved** through openpyxl load→save in this version (tested). A prior
  "drops comments" belief does not apply.
- **No `#REF` is necessary but NOT sufficient** after a structural edit. The real proof is
  calibration against an Excel-recalculated copy. (Confirmed when an HR-column deletion showed
  zero #REF but could only be trusted after an Excel round-trip + zero-diff calibrate.)

---

## 9. FIELD NAMING CONVENTIONS

- Lowercase, underscores, no special characters
- `field_name` alone = scalar or vector
- `field_name` + `index` = one row in a matrix
- `_actual` suffix = historical value (e.g. `mgmt_fee_actual`)
- Index examples: `bermuda_pb`, `austin`, `turks`, `fritholme`, `elbow`, `lsat`, `l4`
- Person codes: initials lowercase e.g. `ce`, `ws`, `cl`, `rf`, `mt`

---

## 10. PYTHON MODEL ARCHITECTURE

### Key functions (reusable across all sheets)

```python
def find_section_boundaries(ws):
    """Scan row 3 for 'actual' and 'i/o' markers."""
    # Returns: (actual_start_col, io_start_col)

def get_opening_col(ws, projections_from):
    """Last actuals col with date strictly < projections_from."""

def get_period_cols(ws, from_date, to_date):
    """Model i/o cols between from_date and to_date."""
    # Returns: {'Q2_2026': 17, 'Q3_2026': 18, ...}

def get_actuals_col(ws, period_label):
    """Find actual col for a period — in actuals zone only."""

def read_sheet(ws, cols, io_col=2, field_col=3, index_col=4):
    """Read all rows by field_name+index into dict."""

def read_c12_inputs(wb):
    """Read scalar assumptions + date controls from c12_inputs."""

def write_outputs(ws, results, period_cols, actuals_through):
    """Write o row results — skips if actual exists."""
```

### .c12_hr sheets
Python does NOT replicate HR salary calculations.
Only reads four summary `o` rows per department:
- `hr_grand_total`
- `hr_reimbursable`
- `hr_not_reimbursable`
- `hr_only_c12`

Excel calculates these from salary/allocation data in `.c12_##_hr` sheets.
Python treats them as inputs.

### Actuals override logic
```python
# At write time:
actual_col = get_actuals_col(ws, period)
if actual_col:
    actual_val = ws.cell(r, actual_col).value
    if isinstance(actual_val, (int, float)):
        continue  # actual exists — don't overwrite
```

---

## 11. C12 MODEL STATUS — COMPLETE & VALIDATED

### Built and calibrating to ZERO
- `c12_dept.py` — shared engine: `load`, `compute_dept` (pure), `diff_o_rows`, `calibrate`,
  and all generic readers (find_section_boundaries, get_opening_col, get_period_cols,
  read_sheet, read_c12_inputs, read_hr_outputs, section totals, etc.).
- `c12_hm.py`, `c12_ds.py`, `c12_bif.py` — thin dept wrappers (SHEET/HR config only).
  (Note filenames avoid `&`: `c12_ds.py`, `c12_bif.py`; sheet names keep it: `c12_d&s`, `c12_bi&F`.)
- `c12_oc.py` — full consolidation (`compute_oc` pure, takes in-memory dept_results; thin
  `calibrate`). ~9 o-rows incl. distribution_required, loc_balance (MIN-capped roll-forward),
  the AR cumulatives.
- `c12_runner.py` — flat orchestrator: one load → 3 depts (pure) → in-memory into compute_oc
  → diff each vs Excel. ALL MODULES OK (zero) on the cache-fresh master and on c12_io.
- The `cumm_accrued_participation` `i`→`o` fix is in place on all three dept pages.

### Validated THREE independent ways
1. **Calibration** — Python reproduces Excel's o-values exactly (zero diffs).
2. **check_model.py** — full master's `lgh_diags` error matrix all zeros.
3. **Backtest** (`backtest.py`) — set projections_from to a period that also has actuals;
   the computed model column reproduces the actual column. (Caveat: the "actuals" currently
   in the file are model-output pasted as values, so this presently proves *internal
   consistency / model-vs-its-own-output*, not model-vs-reality. Becomes a true accuracy test
   once finance enters real actuals.)

### Proven EXTENSIBLE (the architecture is period-agnostic)
- No hardcoded years anywhere in the calculation; one loop body runs per period (verified).
  Same idea as "one column of Excel formulas dragged across" — infinite periods, one logic set.
- Demonstrated by extending the horizon to 2027Q4 with **no code change**: copy i-inputs into
  the new columns, add row-2 dates, bump `projections_to`. Both engines agree on 2027:
  the Excel dragged-formula version and the Python-produced version were diffed cell-by-cell
  across all o-cells = IDENTICAL.

### Production tooling (this session)
- `run_model.py` — CLI entry point: `calibrate` and `produce` modes (section 17).
- `produce.py` — writes o-values into the workbook (projection window only).
- `backtest.py` — model-vs-actual comparison over all overlapping periods.
- `save_clean.py` — the empty-`<v>` corruption fix; used by every Python save.
- `colour_finance.py` / `colour.py` — the two colour schemes.

### Deferred / parked (deliberate)
- `c12_bs` — balance-sheet display page. UNTAGGED (no i/o/c, data starts col B not E), so
  outside Python scope; row 8 and others mispoint (reference c12_oc col E which is empty).
  Decision: leave it; later use it as the worked example for "how to bring a new page into
  the Python model." JP distribution row was moved to `c12_sk` (tagged `i`, field `jp_payment`,
  index `sk`) and is not yet wired into any consumer — part of the future bs build-out.
- Input-surface robustness (row-2 date generation, missing-input guards, finance data-entry
  validation) — parked as "someone else's job"; Excel/data-inputter owns the input surface.

---

## 12. LGH MODEL (future)

### Sheets
`lgh_inputs`, `lgh_returns`, `lgh_ci`, `lgh_cuda`, `lgh_equity`, 
`lgh_pref_equity`, `lgh_debt`, `lgh_other_liabs`, `lgh_other_assets`, `lgh_manco`

### Key complexity
- Debt instruments (27 instruments with principal, rate, term, dates)
- Preferred equity (36 instruments)
- Carried interest calculation
- CUDA schedule
- Opening balances roll forward each quarter

### Architecture
Same i/o/c framework as c12.
`.` pages for debt/pref instrument details (stay in Excel for finance team).
Python reads summary outputs only.

---

## 13. GIT WORKFLOW

```bash
git add .
git commit -m "message"        # new files
git commit -am "message"       # existing files only
git push origin main           # push to GitHub
git pull origin main           # get changes
git log --oneline              # view history
```

---

## 14. AGENTS (from metacode)

| Code | Name | Role |
|------|------|------|
| HoDME | Ryan | Head of Development, Modelling & Execution |
| HoBIF | Charity | Head of Business Infrastructure & Finance |
| HoHM | Chris | Head of Hospitality Management |
| CFO | Zandra | Chief Financial Officer |
| FA | Terry | Financial Accountant |
| MA | Daniel | Management Accountant |
| IM | Thomas | Investment Manager |
| FO | Danielle | Family Office |
| FM | Kevin | Financial Modeller |
| TM | TBD | Technology Manager |
| Prog | TBD | Programmer |
| BA | Pilar | Business Administrator |
| LC | TBD | Legal Counsel |

---

## 15. KEY FILES IN REPO

| File | Purpose |
|------|---------|
| `financial_model_excel_claude.xlsx` | Full master model (36 sheets) — golden reference |
| `c12_io.xlsx` | Reduced C12 production cut (15 sheets) — the working/deliverable file |
| `c12_dept.py` | Shared C12 engine |
| `c12_hm.py` / `c12_ds.py` / `c12_bif.py` | Department wrappers |
| `c12_oc.py` | Consolidation |
| `c12_runner.py` | In-memory orchestrator (calibrate) |
| `run_model.py` | CLI entry point — calibrate / produce |
| `produce.py` | Writes o-values into the workbook |
| `backtest.py` | Model-vs-actual comparison |
| `save_clean.py` | openpyxl empty-`<v>` corruption fix (used by every save) |
| `colour_finance.py` | Finance colour scheme (green=update / grey=leave) |
| `colour.py` | Diagnostic colour scheme |
| `check_model.py` | Diagnostic — full master error matrix |
| `context_document.md` | This file |
| `C12_MODEL_GUIDE.md` | Finance-facing guide (what each file is + what to update) |
| `metacode.md` | Operational process document |

---

## 16. RUNNING THE MODEL (run_model.py)

From the folder containing the modules and the workbook:

```bash
python run_model.py calibrate c12_io.xlsx
    # compute & COMPARE to the file's o-values; writes nothing. Use after editing inputs.

python run_model.py produce c12_io.xlsx --out c12_io_run.xlsx
    # compute & WRITE o-values into a NEW file (originals safe). This "fills the greys".
    # drop --out to overwrite in place (must close the file in Excel first — else
    # PermissionError: file is locked).
```

After a produce run, OPEN the file in Excel and let it recalc (it will prompt) so diags and
display pages refresh. Caches are intentionally cleared by save_clean; fullCalcOnLoad handles it.

Common gotchas seen in practice: file open in Excel → PermissionError on save (close it);
output written to `--out` name, not the original (look in the right file); o-cells only change
in the **projection window** columns — pre-projection columns stay as formulas.

---

## 17. NEXT: LGH PHASE

Open question from the start of the LGH work: do C12/CS2 need full quarterly build-out to 2030
before LGH, or just "working to 2026Q4"? The extensibility result reframes this: **extending
the horizon is a data operation, not a rebuild** — so C12/CS2 don't need to be pre-built to
2030; they extend on demand. The real first task is to read the **registry** to see exactly
which C12/CS2 cells LGH consumes and at what period granularity (quarterly vs annual out-years);
that determines build order. Then apply the same i/o/c framework + the proven toolchain to the
lgh_* sheets.

---

## 18. IMPORTANT NOTES FOR CLAUDE

1. Never use xlwings — use openpyxl and pandas only.
2. Always load two workbooks: one normal (for writing), one data_only (for reading).
3. Column detection is always dynamic via row 3 markers — never hardcode column numbers.
4. Python never touches `.c12_##_hr` internals — only reads four summary o rows.
5. The `projections_from` / `projections_to` dates in c12_inputs drive what Python calculates.
6. field_name + index = unique key — always look up by name, never row number.
7. **Every Python save of a workbook goes through `save_clean`** (empty-`<v>` fix), or Excel
   throws the repair prompt and may delete formulas.
8. **Calibrate/backtest only against an Excel-saved file** (fresh caches). An all-`excel=0.00`
   diff wall is the cache-strip false alarm, not breakage — recalc in Excel and re-check.
9. An `i`-row may hold a formula; read its VALUE, never replicate the formula.
10. To drag a formula to new period columns, use `openpyxl.formula.translate.Translator`.
11. Structural column/row edits don't auto-translate refs in openpyxl — verify, or do in Excel,
    and confirm with a post-recalc calibration (no #REF is not enough).
12. Growth comes only from inputs; the engine has no built-in escalation.
13. Memory + past-chat search persist across conversations; this document is the durable
    long-term memory. Procedures live in the committed scripts — run them, don't re-derive.
