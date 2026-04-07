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

### Graphical UI (Windows-friendly)

From this folder:

```bash
python export_segments_gui.py
```

Or double-click **`Run_Export_GUI.bat`** (adds `.docx` files or a whole folder, pick output `.xlsx`, then export).

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

