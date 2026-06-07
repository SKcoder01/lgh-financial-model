"""
run_model.py — command-line entry point for the C12 model.

USAGE (from the folder containing the model files and your workbook):

  python run_model.py calibrate c12_io.xlsx
        Compute the model and COMPARE against the workbook's o-values.
        Writes nothing. Prints OK / diffs per module. Use this to check the
        model still reproduces the workbook after you change inputs.

  python run_model.py produce c12_io.xlsx
        Compute the model and WRITE the results into the o-cells (replacing
        their formulas with numbers) for the projection window. Saves the file
        in place. THIS is the run that "fills the greys".

  python run_model.py produce c12_io.xlsx --out c12_io_run.xlsx
        Same, but write to a new file instead of overwriting the input.

Notes:
- After a produce run, OPEN the file in Excel and let it recalculate (it will
  prompt briefly) so c12_diags and any display pages refresh.
- Re-running calibrate after editing inputs should still say OK if your edits
  were to i-cells the model reads; if you edited something the model computes,
  you'll see a diff — that's the model disagreeing with a stale o-value.
"""
import sys, os, datetime
import c12_dept as cd
import c12_oc
import c12_runner

def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    mode = sys.argv[1].lower()
    path = sys.argv[2]
    out  = path
    if '--out' in sys.argv:
        out = sys.argv[sys.argv.index('--out') + 1]
    if not os.path.exists(path):
        print(f"ERROR: file not found: {path}"); sys.exit(1)

    cd.EXCEL_FILE = path

    if mode == 'calibrate':
        print(f"CALIBRATE (no writes): {path}\n")
        c12_runner.run(path, verbose=True)

    elif mode == 'produce':
        import produce
        print(f"PRODUCE (writes o-values): {path} -> {out}\n")
        # safety: calibrate first so we only overwrite when model agrees with file
        print("Pre-check (calibrate):")
        c12_runner.run(path, verbose=True)
        print("\nWriting results:")
        n = produce.produce(path, out, verbose=True)
        print(f"\nDONE. Open {out} in Excel and let it recalculate.")
    else:
        print(f"Unknown mode '{mode}'. Use 'calibrate' or 'produce'.")
        print(__doc__); sys.exit(1)

if __name__ == '__main__':
    main()
