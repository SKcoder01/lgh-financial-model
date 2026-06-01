"""c12_ds.py — Design & Structuring. Thin config over the shared c12_dept engine.
   Operates on sheet 'c12_d&s' (the '&' is kept in the SHEET name, not the filename,
   so the module imports/runs cleanly from a shell)."""
import c12_dept
SHEET, HR = 'c12_d&s', '.c12_d&s_hr'

def calibrate(excel_file=c12_dept.EXCEL_FILE, **kw):
    return c12_dept.calibrate(SHEET, HR, excel_file, **kw)

if __name__ == '__main__':
    calibrate()
