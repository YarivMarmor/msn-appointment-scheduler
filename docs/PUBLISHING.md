# Maintainer publishing checklist

## Before making the repository public

- [ ] Inspect original `.xlsm` workbook for patient names, identifiers, sensitive data, hidden sheets, connections, comments, embedded objects, author metadata, and cached outputs. Remove or anonymize them if present, using an *approved separate copy*; the bundled workbook here is unchanged.
- [ ] Confirm you own or may redistribute the Excel workbook and embedded VBA, sample data, and graphics.
- [ ] Run Excel and Python under identical inputs and compare outputs statistically.
- [ ] Choose an explicit license (for example MIT, BSD-3-Clause, Apache-2.0, or a more restrictive license as appropriate). Do not add one without approval.
- [ ] Decide on proper authorship, affiliation, and citation wording.
- [ ] Build and test the EXE on a clean Windows machine, verify it runs without Python installed.
- [ ] Consider code signing the EXE and, separately, VBA macro signing (requires an appropriate signing certificate).

## Upload through GitHub's web interface

1. Sign in at https://github.com/new.
2. Repository name suggestion: `msn-appointment-scheduler`.
3. Choose **Private** initially, until the checks above are completed. Do not create an additional README or `.gitignore` at creation.
4. Upload this repository's contents (not the containing outer folder) with **Add file -> Upload files**, or use Git (recommended for future development).
5. In the **Actions** tab, run **Build Windows EXE** -> **Run workflow**. Download its `MSN_Appointment_Scheduler_windows_unsigned` artifact, unzip it, and test.
6. After validation, go to **Releases -> Draft a new release**, e.g. tag `v0.1.0-preview`, attach the EXE and include a caution about its unsigned status and preliminary validation.
7. Switch repository to **Public** only after you have reviewed the workbook and rights to publish it.

## Git command-line alternative

From inside the repository folder, if Git is installed:

```powershell
git init
git add .
git commit -m "Initial Python and legacy VBA release"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/msn-appointment-scheduler.git
git push -u origin main
```

Replace `YOUR_USERNAME`, and create the empty GitHub repository first. GitHub requires authentication to push.
