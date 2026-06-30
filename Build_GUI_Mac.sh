#!/usr/bin/env bash
# ============================================================
#  Build_GUI_Mac.sh  —  builds dist/InterviewCodingToExcel.app
#  Run ONCE on a Mac with Python 3; share the .app (or a zip of it).
#
#  To run with Python instead: ./Run_Export_GUI.command
# ============================================================

set -euo pipefail
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo ""
  echo "python3 not found."
  echo "Install from https://www.python.org/downloads/ or: brew install python3"
  echo ""
  exit 1
fi

echo ""
echo "Using: $(command -v python3) ($(python3 --version))"
echo "Installing openpyxl and PyInstaller..."
echo ""

python3 -m pip install -r requirements.txt -r requirements-build.txt

echo ""
echo "Building InterviewCodingToExcel.app (may take 1–2 minutes)..."
echo ""

python3 -m PyInstaller \
  --noconfirm \
  --windowed \
  --name "InterviewCodingToExcel" \
  --hidden-import openpyxl \
  --hidden-import tkinter \
  export_segments_gui.py

APP_PATH="dist/InterviewCodingToExcel.app"
if [[ ! -d "$APP_PATH" ]]; then
  echo ""
  echo "Build finished but $APP_PATH was not found."
  exit 1
fi

echo ""
echo "============================================================"
echo " SUCCESS"
echo " App: $(pwd)/$APP_PATH"
echo ""
echo " Share the whole .app folder (or zip it) with colleagues."
echo " First open on another Mac: right-click → Open"
echo "   (unsigned apps are blocked by Gatekeeper otherwise)."
echo "============================================================"
echo ""
open "dist" 2>/dev/null || true
