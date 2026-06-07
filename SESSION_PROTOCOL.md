# SESSION PROTOCOL — LGH/C12 Model
## Version 1.0 — June 2026
## Purpose: ensure every session starts on the current files and ends with all work saved to Git.

These are the steps to run in PowerShell from the repo folder. SK runs them for now;
Claude lists/verifies. A "session end" is whatever SK declares — so if SK is ending a
session, there is by definition work to commit.

**Repo folder:** `C:\Users\steph\Desktop\lgh-financial-model`
**Canonical model file:** `c12_io.xlsx` (single source of truth; currently extends to 2027)
**GitHub:** https://github.com/SKcoder01/lgh-financial-model.git

---

## START-OF-SESSION PROTOCOL

Run these before doing any work, so you're building on the current files — not a stale copy.

### 1. Go to the repo folder (by full path — avoids the "ghost directory" trap)
```powershell
cd C:\Users\steph\Desktop\lgh-financial-model
pwd
```
`pwd` must show `...\lgh-financial-model` with no subfolder. If you renamed/moved a folder
while a shell was inside it, `cd ..` breaks — always `cd` by full path to recover.

### 2. Pull the latest from GitHub (in case anything changed elsewhere)
```powershell
git pull origin main
```
If it says "Already up to date," good. If it pulls changes, you were behind — now current.

### 3. Confirm the canonical file and scripts are present at root
```powershell
ls run_model.py, c12_io.xlsx, c12_dept.py, c12_oc.py, c12_runner.py
```
All five must list. If any is missing you're in the wrong folder or a copy drifted into a
subfolder — fix before proceeding. Run scripts ONLY from the folder that holds c12_io.xlsx.

### 4. Smoke-test: does the model still calibrate against the canonical file?
```powershell
python run_model.py calibrate c12_io.xlsx
```
Expect four "OK zero diffs" lines + "ALL MODULES OK".
- If you get a "FAIL" wall where every diff is `excel=0.00` vs a non-zero number, that is the
  **cache false-alarm**: c12_io.xlsx was last saved by Python, not Excel. Open it in Excel,
  let it recalc, save, close, and re-run. Then it should be zero.
- If you get a real non-zero diff, the file and the code disagree — investigate before working.
- If it errors on import, a module is missing from the folder (e.g. c12_bif.py).

### 5. Note where we left off
Open `context_document.md` (section 11 = status, section 17 = next steps). That's the
breadcrumb for what to do this session.

---

## END-OF-SESSION PROTOCOL

Run these when SK declares the session over, to save everything to Git.

### 1. Close the workbook in Excel
Excel locks open files — a Python save or git operation can fail or grab a half-written file.
Close `c12_io.xlsx` (and any other open workbook) in Excel first.

### 2. Final calibrate against the canonical file (regression check)
```powershell
python run_model.py calibrate c12_io.xlsx
```
Confirm "ALL MODULES OK" before saving the session. If it's not zero, decide whether that's
intended (you changed the model) and update the diags baseline, or a bug to fix first.
(Same cache false-alarm caveat as start-step 4 applies — recalc in Excel if needed.)

### 3. Move interim / derived files to archive-dump
Single canonical file rule: `c12_io.xlsx` stays; derived and experiment files get archived.
Derived = anything regenerable from c12_io.xlsx + scripts (e.g. `c12_io_run.xlsx`,
`*_march_test.xlsx`, `*_dragM.xlsx`, `*_ext20xx.xlsx`). SK's own deliberately-kept files stay.
```powershell
# example — move the regenerable outputs out of the tracked folder
Move-Item c12_io_run.xlsx, c12_io_march_test.xlsx archive-dump\ -Force
```
(`archive-dump` is gitignored — see end-step 5 check.)

### 4. Delete throwaway caches (regenerate automatically)
```powershell
Remove-Item -Recurse -Force __pycache__ -ErrorAction SilentlyContinue
```

### 5. Review what Git will commit — the gate step
```powershell
git status
```
You SHOULD see: your real changes (edited scripts, context doc, c12_io.xlsx if changed,
new files you made). You should NOT see: `archive-dump/`, `__pycache__/`, `~$*.xlsx`
(Excel lock files), or anything surprising. If something unexpected appears, stop and look
before committing. (Identical-content files won't appear — Git tracks content, not dates.)

Confirm `.gitignore` covers the ignore-list:
```powershell
Get-Content .gitignore
```
Should contain at least: `archive-dump/`  `__pycache__/`  `*.pyc`  `~$*.xlsx`  `~$*.xlsm`  `*.tmp`
(If you renamed `old/` to `archive-dump/`, make sure the .gitignore line matches the new name.)

### 6. Stage, commit, push
```powershell
git add -A
git commit -m "Session <date>: <one-line summary of what changed>"
git push origin main
```
If push is rejected as "behind", run `git pull origin main`, resolve, then push again.

### 7. Confirm it landed on GitHub
```powershell
git status
```
Should say "Your branch is up to date with 'origin/main'." That means the session is saved.

### 8. Update the breadcrumb (if state changed materially)
If the project state moved (new module, new validation, a phase completed), update
`context_document.md` (status section + next steps) so the NEXT session starts informed.
This document is the durable memory across conversations — keep it current.

---

## QUICK REFERENCE — the failure modes these guard against
- Wrong directory / ghost folder after a rename → start-step 1 (`cd` full path + `pwd`).
- Stale files / behind remote → start-step 2 (`git pull`), end-steps 6–7 (push + confirm).
- Cache false-alarm FAIL wall → recalc in Excel; never calibrate a Python-saved file.
- Excel lock blocking save → end-step 1 (close Excel).
- Duplicate module copies in subfolders → start-step 3 (run only from the canonical folder).
- Committing junk (archive, caches, lock files) → end-step 5 (`git status` gate + `.gitignore`).
- Losing the canonical file among experiments → single-canonical rule, end-step 3 (archive the rest).
