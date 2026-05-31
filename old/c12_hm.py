"""c12_hm.py — Hotel Management. Thin config over the shared c12_dept engine."""
import c12_dept
SHEET, HR = 'c12_hm', '.c12_hm_hr'

def calibrate(excel_file=c12_dept.EXCEL_FILE, **kw):
    return c12_dept.calibrate(SHEET, HR, excel_file, **kw)

if __name__ == '__main__':
    calibrate()
