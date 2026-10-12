MSN Appointment Scheduler
Appointment scheduling tool with two editions:
Edition	Download / Run	Requirements
Python desktop application	Windows executable (when published under Releases), or run from source	EXE: Windows; source: Python 3.11+
Original Excel/VBA workbook	`legacy/Appointment_Scheduler_MSN_Tool.xlsm`	Desktop Microsoft Excel on Windows with macros permitted
The scheduler uses Monte Carlo simulation to construct appointment times under a user-selected quality-of-service target. Parameters include service-time distributions (normal, uniform, empirical), multiple servers, no-show probability, and acceptable waiting time.
> **Release status: preview / validation required.** The Python port has basic automated tests, but a systematic cross-check against VBA outputs has **not** been completed. Do not rely on its schedules for real clinical operations without independent validation.
Research foundation and citation
This software implements the multi-server numerical-based (MSN) appointment-scheduling approach described in:
> Shnits, B., Bendavid, I., & Marmor, Y. N. (2020). **An appointment scheduling policy for healthcare systems with parallel servers and pre-determined quality of service.** *Omega*, **97**, Article 102095. https://doi.org/10.1016/j.omega.2019.08.002
The study addresses appointment scheduling with parallel servers, stochastic service times, possible no-shows, and a pre-specified quality-of-service requirement. The Excel/VBA workbook is the original implementation provided in this repository; the Python desktop application is a port under validation. The research article defines the scientific method; the present repository provides software implementations and does not replace the published article.
When using the scheduling method in academic work, please cite the paper above. If you use or modify this software, also cite the repository using GitHub's Cite this repository feature (configured by `CITATION.cff`).
For end users
Python edition: When a Windows build has been published, visit the GitHub repository's Releases page and download `MSN_Appointment_Scheduler.exe`. It does not require Excel or a Python installation. An unsigned EXE may prompt Windows SmartScreen; verify the source and checksum before choosing to run it. No ready-made EXE is included in this source repository.
Excel edition: Download the workbook using the link above (use GitHub's Download raw file option). Read How to run unsigned Excel macros safely before opening it. Never enable macros for an unfamiliar file.
Python developer setup
Use a fresh virtual environment (recommended, especially with Anaconda installed):
```powershell
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -e ".[dev]"
.venv\Scripts\python.exe run_app.py
```
Alternatively, on a configured Python environment: `python -m pip install -e .` then `python run_app.py`.
If importing SciPy fails with a NumPy ABI error, do not downgrade packages in your Anaconda base environment. Create a new virtual environment as above. `numpy>=1.26,<2` is specified for compatibility with the dependencies tested during migration.
Build Windows EXE
On Windows, run `build_exe.bat` in a project environment with Python installed. Output: `dist\MSN_Appointment_Scheduler.exe`. Alternatively, run the Build Windows EXE workflow in the GitHub Actions tab; download the ZIP artifact from that workflow run. The generated EXE is not digitally signed. A maintainer must inspect and validate the artifact before uploading it to a GitHub Release.
Repository layout
```text
src/appointment_scheduler/    Python simulation and Tkinter GUI
run_app.py                   Application launcher
legacy/                      Original Excel workbook and extracted VBA text
labs/                        (reserved for validation notebooks)
tests/                       Python automated tests
docs/EXCEL_MACROS.md         Unsigned macro instructions
.github/workflows/           Continuous tests and manual Windows build
build_exe.bat                Windows executable build helper
```
Known differences and limitations
Compatibility with the VBA implementation requires scenario-by-scenario quantitative validation.
The empirical distribution retains a possible legacy off-by-one interpolation issue pending a documented decision.
The original VBA workbook is provided as received, including historical results or metadata it may contain. Review it before making your GitHub repository public.
This is an analytical tool, not a clinically validated medical device or scheduling system.
Contributions, attribution and license
Please open issues for test scenarios, bugs, or feature requests. The repository does not yet declare an open-source license; until the owner chooses one, uploading source code publicly does not grant others broad permission to reuse or redistribute it. Add `LICENSE` with the copyright owner's approval before advertising the project as open source.
