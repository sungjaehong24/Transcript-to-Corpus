#!/usr/bin/env bash
# Double-click in Finder (after: chmod +x Run_Export_GUI.command once, or run from Terminal).
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  osascript -e 'display alert "Python 3 not found" message "Install Python 3 from python.org or run: brew install python3"' 2>/dev/null || echo "Python 3 not found."
  read -r -p "Press Enter to close…"
  exit 1
fi

python3 -m pip install -q -r requirements.txt 2>/dev/null || true
python3 export_segments_gui.py
status=$?
if [[ $status -ne 0 ]]; then
  read -r -p "Press Enter to close…"
fi
exit $status
