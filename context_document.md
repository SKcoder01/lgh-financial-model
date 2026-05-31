# LGH/C12 Financial Model — Claude Context Document
## Version: 1.0 — June 2026
## Purpose: Paste at start of new Claude conversation to restore full context

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

---

## 6. COLUMN STRUCTURE (all c12 sheets)

### Row 3 Markers (critical for Python)
- `"actual"` — appears in first column of actuals section
- `"i/o"` — appears in first column of model i/o section
- Python scans row 3 to find these markers dynamically

### Column Zones
- **Col A** — row label
- **Cols B-D** — i/o/c, field_name, index (metadata)
- **Col E** — Q4 2024 actual
- **Cols F-K** — Q1 2025 to Q2 2026 actuals (E-K = actual zone)
- **Cols L onwards** — model i/o projection columns (i/o zone)

### Date Controls (in c12_inputs)
- `actuals_thru` — last period with actual data (date)
- `projections_from` — first period Python calculates (date)
- `projections_to` — last period Python calculates (date, currently 2026-12-31)

**Opening balance** = last actuals col with date strictly < projections_from

---

## 7. COLOUR CODING

Applied to all sheets:

| Colour | Meaning |
|--------|---------|
| 🟢 Dark green | Pure hardcoded number — `i` row input |
| ⬜ Grey | `o` row — Python writes result back here |
| ⬜ White | `c` row or pure calculation formula |
| 🟠 Orange | IF formula (actual vs model logic) |
| 🔵 Light blue | Date, year 2024-2030, Q1-Q4 label, or actual column value |
| 🟡 Yellow | Text label/string |
| 🟨 Light yellow | Formula referencing a text cell |
| ⬜ Light grey | SUM formula |
| 🟫 Light brown | Header rows (rows 2-3) |
| **Bold + thick border** | Cell that registry reads from |

**Run colour coding:** upload file and ask Claude to "run excel diags and colour coding"

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
When model changes intentionally: load file data_only, copy live cols C-L to baseline cols N-V in lgh_diags.

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

## 11. C12 MODEL STATUS

### Completed
- `c12_hm.py` — Hotel Management ✓ Zero differences vs Excel
- All c12 sheets have i/o/c markers, field_name, index in cols B-D
- `.c12_alloc` sheet created with allocation matrix
- `.c12_##_hr` sheets: salary/bonus inputs marked `i`, four summary rows marked `o`
- Row 3 markers (`actual` / `i/o`) on all c12 sheets
- Dynamic column detection — no hardcoded column numbers

### To Build
- `c12_d&s.py` — Design & Structuring (same pattern as c12_hm)
- `c12_bi&f.py` — Business Infrastructure & Finance (same pattern)
- `c12_oc.py` — Consolidated P&L
- `c12_bs.py` — Balance Sheet
- `c12_runner.py` — runs all c12 modules in sequence

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
| `financial_model_excel_claude.xlsx` | Main Excel model |
| `check_model.py` | Diagnostic — checks error matrix |
| `c12_hm.py` | Hotel Management Python model |
| `context_document.md` | This file |
| `metacode.md` | Operational process document |

---

## 16. IMPORTANT NOTES FOR CLAUDE

1. Never use xlwings — use openpyxl and pandas only
2. Always load two workbooks: one normal (for writing), one data_only (for reading calculated values)
3. Column detection is always dynamic via row 3 markers — never hardcode column numbers
4. Python never touches `.c12_##_hr` internals — only reads four summary o rows
5. The `projections_from` date in c12_inputs drives what Python calculates
6. When asked to "run excel diags" — do all four checks listed in section 8
7. When updating baseline — copy live model cols to baseline cols in lgh_diags
8. field_name + index = unique key — Python always looks up by name not row number

