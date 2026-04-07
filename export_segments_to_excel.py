"""
Export Word (.docx) comment-based coding to Excel.

- One row per interview segment (same anchored body text).
- One column per coder (Word comment w:author); different coders are never merged in one cell.
- Multiple codes from the same coder on the same segment are joined with '; ' in that cell only.
- Blank cell if a coder did not code that segment.

Requires: openpyxl (pip install openpyxl)
Offline; reads OOXML inside .docx (zip).
"""

from __future__ import annotations

import argparse
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from typing import Dict, List, Tuple

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _local(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def _attr_id(elem: ET.Element) -> str | None:
    """w:id on comment range elements (namespace may vary)."""
    for key, val in elem.attrib.items():
        if key.endswith("}id") or key == "id":
            return val
    return None


def _collect_w_t_text(node: ET.Element) -> str:
    parts: List[str] = []
    for el in node.iter():
        if _local(el.tag) == "t":
            if el.text:
                parts.append(el.text)
            if el.tail:
                parts.append(el.tail)
    return "".join(parts)


def _parse_comments_xml(zf: zipfile.ZipFile) -> Dict[str, Tuple[str, str]]:
    """
    comment_id -> (author, comment_body_text)
    """
    try:
        data = zf.read("word/comments.xml")
    except KeyError:
        return {}
    root = ET.fromstring(data)
    out: Dict[str, Tuple[str, str]] = {}
    for el in root.iter():
        if _local(el.tag) != "comment":
            continue
        cid = _attr_id(el)
        if cid is None:
            continue
        author = el.get(_w("author")) or el.get("author") or ""
        body = _collect_w_t_text(el)
        body = re.sub(r"\s+", " ", body).strip()
        out[cid] = (author, body)
    return out


def _segment_text_for_comment(doc_root: ET.Element, comment_id: str) -> str:
    """Text strictly between commentRangeStart and commentRangeEnd (document order)."""
    start_tag = _w("commentRangeStart")
    end_tag = _w("commentRangeEnd")
    order = list(doc_root.iter())
    start_idx = end_idx = None
    for i, el in enumerate(order):
        if el.tag == start_tag and _attr_id(el) == comment_id:
            start_idx = i
        if el.tag == end_tag and _attr_id(el) == comment_id:
            end_idx = i
    if start_idx is None or end_idx is None or end_idx <= start_idx:
        return ""
    parts: List[str] = []
    for i in range(start_idx + 1, end_idx):
        el = order[i]
        if _local(el.tag) == "t":
            if el.text:
                parts.append(el.text)
            if el.tail:
                parts.append(el.tail)
    return "".join(parts)


def _normalize_segment_text(s: str) -> str:
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _author_key(name: str) -> str:
    """Map empty author to a single placeholder column."""
    n = (name or "").strip()
    return n if n else "(no author)"


def _format_transcript_layout(s: str) -> str:
    """
    Insert a newline after 'P02 mm:ss' or 'Interviewer mm:ss' (transcript-style line break).
    """
    if not s:
        return s
    return re.sub(
        r"((?:P02|Interviewer)\s+\d{1,2}:\d{2})\s*([^\n])",
        r"\1\n\2",
        s,
    )


def extract_file(path: str) -> List[Tuple[str, str, str, str]]:
    """
    Returns list of (comment_id, author, segment_text, code_text) per comment.
    """
    rows: List[Tuple[str, str, str, str]] = []
    with zipfile.ZipFile(path, "r") as zf:
        try:
            doc_xml = zf.read("word/document.xml")
        except KeyError:
            return rows
        doc_root = ET.fromstring(doc_xml)
        comments = _parse_comments_xml(zf)

        for cid in sorted(comments.keys(), key=lambda x: int(x) if x.isdigit() else x):
            author, code_text = comments[cid]
            raw_seg = _segment_text_for_comment(doc_root, cid)
            seg = _normalize_segment_text(raw_seg)
            rows.append((cid, author, seg, code_text))
    return rows


def _write_excel_workbook(all_rows: List[dict], output_path: str) -> None:
    """Each row dict must include `_author_codes`; that key is removed after save."""
    author_columns: List[str] = sorted(
        {a for r in all_rows for a in r["_author_codes"].keys()}
    )

    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Coding"
    cols = ["Segment_ID", "File_Name", "Coding_Segment", *author_columns]
    ws.append(cols)
    for r in all_rows:
        ac = r["_author_codes"]
        row = [r["Segment_ID"], r["File_Name"], r["Coding_Segment"]]
        row.extend(ac.get(name, "") for name in author_columns)
        ws.append(row)

    wrap_top = Alignment(wrap_text=True, vertical="top")
    max_col = 3 + len(author_columns)
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=max_col):
        for cell in row:
            cell.alignment = wrap_top
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 88
    for i, _name in enumerate(author_columns, start=4):
        ws.column_dimensions[get_column_letter(i)].width = 36

    for r in all_rows:
        del r["_author_codes"]

    out_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    wb.save(out_path)


def export_docx_paths_to_xlsx(
    paths: List[str], output_path: str
) -> tuple[int, List[str]]:
    """
    Process .docx files and write one .xlsx.

    Returns:
        (number of rows written, basenames skipped because they had no comments)

    Raises:
        ValueError: nothing to export
        ImportError: openpyxl not installed
    """
    try:
        import openpyxl  # noqa: F401
    except ImportError as e:
        raise ImportError(
            "Install openpyxl: python -m pip install openpyxl"
        ) from e

    all_rows: List[dict] = []
    skipped: List[str] = []
    seg_id = 1
    for p in paths:
        p = os.path.abspath(p)
        if not os.path.isfile(p):
            continue
        label = os.path.basename(p)
        per = extract_file(p)
        if not per:
            skipped.append(label)
            continue
        for block in build_segment_rows(label, per):
            block["Segment_ID"] = seg_id
            seg_id += 1
            all_rows.append(block)

    if not all_rows:
        raise ValueError(
            "No data to export. Check that your .docx files contain Word comments."
        )

    _write_excel_workbook(all_rows, output_path)
    return len(all_rows), skipped


def build_segment_rows(
    file_label: str, per_comment: List[Tuple[str, str, str, str]]
) -> List[dict]:
    """
    Group by normalized segment within the file; merge codes with '; ' per author only.
    Different authors become separate Excel columns (not merged in one cell).
    """
    groups: Dict[str, List[Tuple[str, str, str, str]]] = defaultdict(list)
    for row in per_comment:
        cid, author, seg, code = row
        key = seg if seg else f"__empty__:{cid}"
        groups[key].append(row)

    out: List[dict] = []
    for key in sorted(groups.keys(), key=lambda k: (k.startswith("__empty__"), k)):
        items = groups[key]
        items.sort(key=lambda t: int(t[0]) if t[0].isdigit() else t[0])
        seg_text = items[0][2] if not key.startswith("__empty__:") else ""
        seg_text = _format_transcript_layout(seg_text)

        by_author: Dict[str, List[Tuple[str, str]]] = defaultdict(list)
        for cid, author, _seg, code in items:
            by_author[_author_key(author)].append((cid, code))

        author_codes: Dict[str, str] = {}
        for au, pairs in by_author.items():
            pairs.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0])
            merged = "; ".join(c for _, c in pairs if c)
            if merged:
                author_codes[au] = merged

        out.append(
            {
                "File_Name": file_label,
                "Coding_Segment": seg_text,
                "_author_codes": author_codes,
            }
        )
    return out


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Word comments -> Excel: one row per segment, one column per coder (author)."
    )
    ap.add_argument(
        "input",
        nargs="+",
        help="One or more .docx paths, or use --dir for a folder of .docx",
    )
    ap.add_argument(
        "-d",
        "--dir",
        action="store_true",
        help="Treat input as directory(ies): process all *.docx inside",
    )
    ap.add_argument(
        "-o",
        "--output",
        required=True,
        help="Output .xlsx path",
    )
    args = ap.parse_args()

    paths: List[str] = []
    if args.dir:
        for d in args.input:
            if not os.path.isdir(d):
                raise SystemExit(f"Not a directory: {d}")
            for name in sorted(os.listdir(d)):
                if name.lower().endswith(".docx") and not name.startswith("~$"):
                    paths.append(os.path.join(d, name))
    else:
        paths = [os.path.abspath(p) for p in args.input]

    if not paths:
        raise SystemExit("No .docx files to process.")

    out_path = os.path.abspath(args.output)
    try:
        n, skipped = export_docx_paths_to_xlsx(paths, out_path)
    except ValueError as e:
        raise SystemExit(str(e)) from e
    except ImportError as e:
        raise SystemExit(str(e)) from e

    print(f"Wrote {n} row(s) -> {out_path}")
    for s in skipped:
        print(f"[skip] No comments or unreadable: {s}")


if __name__ == "__main__":
    main()
