"""
Tkinter front-end for export_segments_to_excel (only extra runtime dep: openpyxl).

Default export: Excel columns no / quote / code (highlight colours + comment text;
same-quote rows merged in the engine). Optional “coder matrix” mode matches CLI
--coder-matrix.

Run:
  python export_segments_gui.py
Windows shortcut:
  Run_Export_GUI.bat
"""

from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

from export_segments_to_excel import export_docx_paths_to_xlsx


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Interview coding → Excel")
        self.minsize(560, 420)
        self.paths: list[str] = []
        self._coder_matrix_legacy = tk.BooleanVar(value=False)
        self._build()

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 6}

        ttk.Label(
            self,
            text=(
                "Add coded Word (.docx) files, choose export path, Export. Default: columns "
                "no / quote / code from highlight colours + appraisal comments "
                '(see project “List of codes”). Tick “Coder matrix…” for legacy one-column-per-author.'
            ),
            wraplength=520,
        ).pack(anchor="w", **pad)

        ttk.Checkbutton(
            self,
            text="Coder matrix export (legacy: one Excel column per comment author)",
            variable=self._coder_matrix_legacy,
        ).pack(anchor="w", **pad)

        lf = ttk.LabelFrame(self, text="Input files")
        lf.pack(fill="both", expand=True, **pad)

        inner = ttk.Frame(lf)
        inner.pack(fill="both", expand=True, padx=6, pady=6)

        btn_row = ttk.Frame(inner)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Add files…", command=self._add_files).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="Add folder…", command=self._add_folder).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="Remove selected", command=self._remove_sel).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="Clear list", command=self._clear).pack(side="left")

        scroll = ttk.Scrollbar(inner)
        scroll.pack(side="right", fill="y")
        self._list = tk.Listbox(
            inner,
            height=12,
            selectmode=tk.EXTENDED,
            yscrollcommand=scroll.set,
            font=("Consolas", 9),
        )
        self._list.pack(side="left", fill="both", expand=True)
        scroll.config(command=self._list.yview)

        out_fr = ttk.LabelFrame(self, text="Output Excel")
        out_fr.pack(fill="x", **pad)
        of = ttk.Frame(out_fr)
        of.pack(fill="x", padx=6, pady=6)
        self._out = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Desktop", "Coding_Matrix_Result.xlsx")
        )
        ttk.Entry(of, textvariable=self._out).pack(side="left", fill="x", expand=True)
        ttk.Button(of, text="Save as…", command=self._pick_save).pack(
            side="left", padx=(8, 0)
        )

        ttk.Button(self, text="Export to Excel", command=self._run).pack(**pad)

        self.status = ttk.Label(self, text="", foreground="#333")
        self.status.pack(anchor="w", padx=10, pady=(0, 8))

    def _refresh_list(self) -> None:
        self._list.delete(0, tk.END)
        for p in self.paths:
            self._list.insert(tk.END, p)

    def _add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Select Word files",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        for p in files:
            ap = os.path.abspath(p)
            if ap not in self.paths:
                self.paths.append(ap)
        self._refresh_list()
        self._hint()

    def _add_folder(self) -> None:
        d = filedialog.askdirectory(title="Select folder (all .docx files will be added)")
        if not d:
            return
        d = os.path.abspath(d)
        for name in sorted(os.listdir(d)):
            if not name.lower().endswith(".docx") or name.startswith("~$"):
                continue
            p = os.path.join(d, name)
            if p not in self.paths:
                self.paths.append(p)
        self._refresh_list()
        self._hint()

    def _remove_sel(self) -> None:
        sel = list(self._list.curselection())
        if not sel:
            return
        for i in reversed(sel):
            del self.paths[i]
        self._refresh_list()
        self._hint()

    def _clear(self) -> None:
        self.paths.clear()
        self._refresh_list()
        self._hint()

    def _pick_save(self) -> None:
        p = filedialog.asksaveasfilename(
            title="Save Excel as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile="Coding_Matrix_Result.xlsx",
        )
        if p:
            self._out.set(p)

    def _hint(self) -> None:
        n = len(self.paths)
        self.status.config(text=f"Ready — {n} file(s) in list")

    def _run(self) -> None:
        if not self.paths:
            messagebox.showwarning("No input", "Add at least one Word (.docx) file.")
            return
        out = self._out.get().strip()
        if not out:
            messagebox.showwarning("Output path", "Choose where to save the .xlsx file.")
            return
        out = os.path.abspath(out)
        self.status.config(text="Working…")
        self.update_idletasks()
        try:
            n, skipped = export_docx_paths_to_xlsx(
                self.paths,
                out,
                coder_matrix=self._coder_matrix_legacy.get(),
            )
        except ValueError as e:
            self.status.config(text="")
            messagebox.showerror("Export failed", str(e))
            return
        except ImportError as e:
            self.status.config(text="")
            messagebox.showerror("Missing package", str(e))
            return
        except OSError as e:
            self.status.config(text="")
            messagebox.showerror("File error", str(e))
            return

        msg = f"Saved {n} row(s).\n\n{out}"
        if skipped:
            msg += "\n\nSkipped (nothing to export):\n• " + "\n• ".join(skipped)
        self.status.config(text=f"Done — {n} row(s) → {os.path.basename(out)}")
        messagebox.showinfo("Done", msg)


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
