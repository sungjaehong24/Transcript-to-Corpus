# Transcript-to-Corpus
Read coded interview transcripts and create coding matrix (i.e., corpus)

# Word interview coding → Excel matrix

Extract **Microsoft Word comments** (anchored to highlighted transcript segments) into a spreadsheet: **one row per coding segment**, **one column per coder** (Word comment author). Multiple codes from the same coder on the same segment are joined with `"; "` (semicolon + space). Cells are left **blank** when a coder did not code that segment.

Runs **offline** (no cloud APIs). Input must be **`.docx`** (not legacy `.doc`).

## Requirements

- **Python 3**
- [`openpyxl`](https://openpyxl.readthedocs.io/)

## Install

```bash
pip install -r requirements.txt
```

## Usage

### Graphical UI

From this folder:

```bash
python3 export_segments_gui.py
```

| OS | Run with Python | Build standalone app (no Python for recipients) |
|----|-----------------|---------------------------------------------------|
| **Windows** | `Run_Export_GUI.bat` | `Build_GUI_EXE.bat` → `dist\InterviewCodingToExcel.exe` |
| **macOS** | `Run_Export_GUI.command` | `Build_GUI_Mac.sh` → `dist/InterviewCodingToExcel.app` |

**Requires Python** on the machine that runs the `.bat` / `.command` or builds the app (`pip install -r requirements.txt` once).

#### Windows — standalone `.exe`

1. On a Windows PC **with Python**, open the project folder.
2. Double-click **`Build_GUI_EXE.bat`** (not `Run_Export_GUI.bat`).
3. Copy **`dist\InterviewCodingToExcel.exe`** to other Windows PCs.

#### macOS — standalone `.app`

PyInstaller **must run on a Mac** (you cannot build a Mac app from Windows/Linux).

1. Copy the project folder to a Mac with **Python 3**.
2. In Terminal:
   ```bash
   cd path/to/Transcript-to-Corpus
   chmod +x Build_GUI_Mac.sh Run_Export_GUI.command
   ./Build_GUI_Mac.sh
   ```
3. Share **`dist/InterviewCodingToExcel.app`** (zip the `.app` if emailing).
4. **First launch on another Mac:** right-click the app → **Open** (unsigned apps are blocked by Gatekeeper if you double-click). Or in Terminal: `xattr -cr InterviewCodingToExcel.app`

If the app window is **blank** (title bar only), rebuild with the latest `main` — older builds missed Tcl/Tk files for the GUI. `Build_GUI_Mac.sh` now uses `interview_coding_mac.spec` with full tkinter bundling.

To run from source on Mac without building: double-click **`Run_Export_GUI.command`** (after `chmod +x` once) or `python3 export_segments_gui.py`.

### Command line

One or more files:

```bash
python export_segments_to_excel.py path/to/interview1.docx path/to/interview2.docx -o output.xlsx
```

All `.docx` in a directory:

```bash
python export_segments_to_excel.py path/to/folder -d -o output.xlsx
```

## How your Word file should be set up

- **Segments** are whatever text sits between each comment’s anchor range in the document body.
- **Codes** are the **comment body text** (what you type in the comment balloon).
- **Coders** are distinguished by Word’s **comment author** (`Review` → display name). Same person, multiple comments on the same anchor → merged in that author’s cell with `"; "`.
- **Modern Word “reply” threads** are still separate comments; grouping is by **anchored segment text** (normalized), not by thread.

Transcript-style line breaks: after `P02 mm:ss` or `Interviewer mm:ss`, the exporter inserts a newline in the **Coding_Segment** cell for readability (Excel wrap is enabled for that column).

## Output columns

| Column | Content |
|--------|---------|
| `Segment_ID` | Running ID across the export |
| `File_Name` | Source `.docx` file name |
| `Coding_Segment` | Anchored segment text |
| *(dynamic)* | One column per **unique comment author** in the batch, with that person’s code(s) for the row |

## Limitations

- **`.docx` only** — extract from OOXML inside the zip package.
- Segments that share **identical normalized text** in the same file may be merged into one row (rare in practice).
- Comment author display names should stay **consistent** across machines if you rely on column headers for coder identity.

## License

Add a `LICENSE` file in your repository if you need one; this project ships without a bundled license unless you add it.

