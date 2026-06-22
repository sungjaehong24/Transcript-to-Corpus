@echo off
REM ============================================================
REM  Build_GUI_EXE.bat  --  makes dist\InterviewCodingToExcel.exe
REM  (Python required ON THIS PC only; recipients need no Python.)
REM
REM  To run the GUI with Python instead, use Run_Export_GUI.bat
REM ============================================================

cd /d "%~dp0"

set "PY="
where python >nul 2>&1 && set "PY=python"
if not defined PY where py >nul 2>&1 && set "PY=py -3"
if not defined PY (
  echo.
  echo Python not found. Install from https://www.python.org/downloads/
  echo Tick "Add python.exe to PATH", then run this file again.
  echo.
  pause
  exit /b 1
)

echo.
echo Using: %PY%
echo Installing packages (openpyxl, PyInstaller)...
echo.

%PY% -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 goto :err

echo.
echo Building standalone .exe (may take 1-2 minutes)...
echo.

%PY% -m PyInstaller ^
  --noconfirm ^
  --onefile ^
  --windowed ^
  --name "InterviewCodingToExcel" ^
  --hidden-import openpyxl ^
  --hidden-import tkinter ^
  export_segments_gui.py

if errorlevel 1 goto :err

if not exist "dist\InterviewCodingToExcel.exe" (
  echo.
  echo Build finished but dist\InterviewCodingToExcel.exe was not found.
  goto :err
)

echo.
echo ============================================================
echo  SUCCESS
echo  File: %~dp0dist\InterviewCodingToExcel.exe
echo  Copy that .exe to any Windows PC (no Python needed there).
echo ============================================================
echo.
start "" explorer "%~dp0dist"
pause
exit /b 0

:err
echo.
echo Build failed. Read the messages above.
echo.
pause
exit /b 1
