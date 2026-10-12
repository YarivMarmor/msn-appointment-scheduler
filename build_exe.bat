@echo off
setlocal
python -m pip install --upgrade pip
python -m pip install -e .[dev]
pyinstaller --noconfirm --clean --onefile --windowed --paths src --name MSN_Appointment_Scheduler run_app.py

echo.
echo EXE created under dist\MSN_Appointment_Scheduler.exe
endlocal
