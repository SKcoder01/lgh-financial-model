# LGH/C12 Model — Colour Coding Context
## For pasting into a new Claude conversation

---

## TASK
Apply colour coding to the LGH/C12 Excel financial model using openpyxl in Python.

---

## COLOUR SCHEME

| Colour | Hex | Meaning |
|--------|-----|---------|
| Dark green | `FF00B050` | Pure hardcoded number — definite input (`i` row) |
| Grey | `FFD3D3D3` | `o` row — Python writes result back here |
| White | `FFFFFFFF` | `c` row or pure calculation formula |
| Orange | `FFFFA500` | IF formula (actual vs model logic) |
| Light blue | `FFADD8E6` | Date, year 2024-2030, Q1-Q4 label, or actual column value |
| Yellow | `FFFFFF00` | Text label/string |
| Light yellow | `FFFFFFCC` | Formula referencing a text cell |
| Light grey | `FFD3D3D3` | SUM formula |
| Light brown | `FFCC9966` | Header rows (rows 2-3) |
| Bold + thick border | — | Cell that registry reads from |

---

## SHEET STRUCTURE

Every c12 sheet has:
- **Col A** — row label (yellow)
- **Col B** — i/o/c marker
- **Col C** — field_name
- **Col D** — index
- **Row 2** — dates (light brown)
- **Row 3** — period labels + section markers `"actual"` and `"i/o"` (light brown)
- **Cols E-K** — actual data columns (light blue if numeric)
- **Cols L onwards** — model i/o projection columns (dark green if `i`, grey if `o`, white if `c`)

---

## ROW 3 MARKERS

Row 3 contains two key markers:
- `"actual"` — appears in the first column of the actuals section (col E)
- `"i/o"` — appears in the first column of the model i/o section (col L)

These markers are used by Python to find column boundaries dynamically.

---

## REGISTRY SOURCES

Cells that the `registry` sheet reads from get:
- Thick border on all four sides
- Bold font

---

## PYTHON COLOUR CODING FUNCTION

```python
import openpyxl, re
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import column_index_from_string
from datetime import datetime, date

def apply_colour_coding(wb, sheet_names_to_colour=None):
    """
    Apply colour coding to all sheets (or specified sheets).
    Uses registry sheet to identify cells with thick borders.
    """
    WHITE        = PatternFill(fill_type='solid', fgColor='FFFFFFFF')
    DARK_GREEN   = PatternFill(fill_type='solid', fgColor='FF00B050')
    GREY         = PatternFill(fill_type='solid', fgColor='FFD3D3D3')
    YELLOW       = PatternFill(fill_type='solid', fgColor='FFFFFF00')
    LIGHT_BROWN  = PatternFill(fill_type='solid', fgColor='FFCC9966')
    LIGHT_BLUE   = PatternFill(fill_type='solid', fgColor='FFADD8E6')
    ORANGE       = PatternFill(fill_type='solid', fgColor='FFFFA500')
    THICK        = Side(style='thick')
    THICK_BORDER = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
    NO_BORDER    = Border()

    # Find registry sheet
    reg_sheet = next((s for s in wb.sheetnames if s.lower().startswith('registry')), None)

    def get_registry_sources(target_sheet):
        sources = set()
        if not reg_sheet: return sources
        ws_reg = wb[reg_sheet]
        for row in ws_reg.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith('='): continue
                for m in re.finditer(r"'?([A-Za-z0-9_&\s\[\].]+)'?!\$?([A-Z]+)\$?(\d+)", cell.value):
                    if m.group(1).strip() == target_sheet:
                        sources.add((column_index_from_string(m.group(2)), int(m.group(3))))
        return sources

    def find_section_boundaries(ws):
        """Scan row 3 for 'actual' and 'i/o' markers."""
        actual_start = io_start = None
        for col in range(1, ws.max_column + 1):
            v = ws.cell(3, col).value
            if isinstance(v, str):
                if v.strip().lower() == 'actual' and actual_start is None:
                    actual_start = col
                elif v.strip().lower() == 'i/o' and io_start is None:
                    io_start = col
        return actual_start, io_start

    def is_if_formula(v):
        return isinstance(v, str) and bool(re.match(r'^=\s*IF\(', v, re.IGNORECASE))

    sheets = sheet_names_to_colour or wb.sheetnames

    for sheet_name in sheets:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        registry_sources = get_registry_sources(sheet_name)
        _, io_start = find_section_boundaries(ws)
        actual_limit = io_start if io_start else ws.max_column + 1

        # Build row io map from col B
        row_io = {}
        for r in range(1, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if isinstance(b, str) and b.strip().lower() in ('i', 'o', 'c'):
                row_io[r] = b.strip().lower()

        for r in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(r, col)
                v = cell.value
                sz = cell.font.size if cell.font and cell.font.size else 11
                nm = cell.font.name if cell.font and cell.font.name else 'Calibri'

                is_reg = (col, r) in registry_sources
                cell.border = THICK_BORDER if is_reg else NO_BORDER

                # Determine fill
                if r in (2, 3):
                    fill = LIGHT_BROWN
                elif is_if_formula(v):
                    fill = ORANGE
                elif isinstance(v, str) and not v.startswith('='):
                    fill = YELLOW
                elif col < actual_limit and isinstance(v, (int, float)):
                    fill = LIGHT_BLUE  # actuals zone
                elif col >= actual_limit:
                    io_val = row_io.get(r)
                    if io_val == 'i': fill = DARK_GREEN
                    elif io_val == 'o': fill = GREY
                    else: fill = WHITE
                else:
                    fill = WHITE

                cell.fill = fill
                cell.font = Font(name=nm, size=sz, bold=is_reg, italic=False,
                               underline='none', strike=False, color='FF000000')

        print(f"  {sheet_name}: done")
```

---

## HOW TO RUN

```python
wb = openpyxl.load_workbook('financial_model_excel_claude.xlsx')
apply_colour_coding(wb)
wb.save('financial_model_excel_claude.xlsx')
```

---

## NOTES

1. The `registry` sheet is auto-detected by name prefix
2. Row 3 markers (`actual` / `i/o`) determine which columns are actual vs projection
3. All sheets are processed the same way — no special cases needed
4. The `.c12_alloc` sheet has no time series — treat all numeric data rows as `i` (dark green)
5. `c12_inputs` is scalar — treat all numeric values as dark green regardless of column
6. Always load two workbooks when running Python model: one normal (for writing), one data_only (for reading calculated formula values)
7. Never use xlwings — use openpyxl only

