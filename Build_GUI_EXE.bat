@echo off
REM Build dist\InterviewCodingToExcel.exe (standalone GUI, no Python needed on target PC).
REM Requires Python on THIS machine; run from the repo root.

cd /d "%~dp0"

python -m pip install -q -r requirements.txt -r requirements-build.txt
if errorlevel 1 goto :err

python -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name "InterviewCodingToExcel" ^
  --hidden-import openpyxl ^
  --hidden-import tkinter ^
  export_segments_gui.py

if errorlevel 1 goto :err

echo.
echo Done. Run:  dist\InterviewCodingToExcel.exe
echo.
pause
exit /b 0

:err
echo.
echo PyInstaller build failed. Check Python and error messages above.
pause
exit /b 1
