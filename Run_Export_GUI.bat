@echo off
cd /d "%~dp0"
python export_segments_gui.py
if errorlevel 1 pause
