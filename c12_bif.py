"""c12_bif.py — Business Infrastructure & Finance. Thin config over the shared c12_dept engine.
   Operates on sheet 'c12_bi&F' with HR sheet '.c12_bi&f_hr' (note the F/f case difference)."""
import c12_dept
SHEET, HR = 'c12_bi&F', '.c12_bi&f_hr'

def calibrate(excel_file=c12_dept.EXCEL_FILE, **kw):
    return c12_dept.calibrate(SHEET, HR, excel_file, **kw)

if __name__ == '__main__':
    calibrate()
