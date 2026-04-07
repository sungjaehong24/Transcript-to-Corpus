@echo off
REM Run the Interview coding -> Excel GUI from this folder (Python must be on PATH).
cd /d "%~dp0"
python export_segments_gui.py
if errorlevel 1 (
  echo.
  echo Something went wrong. Is Python installed? Try: python --version
  pause
)
