"""
Word 코딩 추출 — 간단 GUI (tkinter, 별도 패키지 없음).

실행: 이 폴더에서
  python export_segments_gui.py
또는 Run_Export_GUI.bat 더블클릭
"""

from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# 같은 폴더의 모듈 import (스크립트 직접 실행 시)
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
        self._build()

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 6}

        ttk.Label(
            self,
            text="처리할 Word 파일(.docx)을 아래 목록에 넣은 뒤, 저장할 엑셀 위치를 정하고 실행하세요.",
            wraplength=520,
        ).pack(anchor="w", **pad)

        lf = ttk.LabelFrame(self, text="입력 파일")
        lf.pack(fill="both", expand=True, **pad)

        inner = ttk.Frame(lf)
        inner.pack(fill="both", expand=True, padx=6, pady=6)

        btn_row = ttk.Frame(inner)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="파일 추가…", command=self._add_files).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="폴더에서 모두 추가…", command=self._add_folder).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="선택 항목 제거", command=self._remove_sel).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="목록 비우기", command=self._clear).pack(side="left")

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

        out_fr = ttk.LabelFrame(self, text="저장할 엑셀")
        out_fr.pack(fill="x", **pad)
        of = ttk.Frame(out_fr)
        of.pack(fill="x", padx=6, pady=6)
        self._out = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Desktop", "Coding_Matrix_Result.xlsx")
        )
        ttk.Entry(of, textvariable=self._out).pack(side="left", fill="x", expand=True)
        ttk.Button(of, text="다른 이름으로…", command=self._pick_save).pack(
            side="left", padx=(8, 0)
        )

        ttk.Button(self, text="엑셀 만들기", command=self._run).pack(**pad)

        self.status = ttk.Label(self, text="", foreground="#333")
        self.status.pack(anchor="w", padx=10, pady=(0, 8))

    def _refresh_list(self) -> None:
        self._list.delete(0, tk.END)
        for p in self.paths:
            self._list.insert(tk.END, p)

    def _add_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="Word 파일 선택",
            filetypes=[("Word 문서", "*.docx"), ("모든 파일", "*.*")],
        )
        for p in files:
            ap = os.path.abspath(p)
            if ap not in self.paths:
                self.paths.append(ap)
        self._refresh_list()
        self._hint()

    def _add_folder(self) -> None:
        d = filedialog.askdirectory(title="폴더 선택 (.docx 모두 포함)")
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
            title="엑셀 저장 위치",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Coding_Matrix_Result.xlsx",
        )
        if p:
            self._out.set(p)

    def _hint(self) -> None:
        n = len(self.paths)
        self.status.config(text=f"대기 중 — 입력 파일 {n}개")

    def _run(self) -> None:
        if not self.paths:
            messagebox.showwarning("입력 없음", "Word 파일을 하나 이상 추가하세요.")
            return
        out = self._out.get().strip()
        if not out:
            messagebox.showwarning("저장 위치", "저장할 .xlsx 경로를 지정하세요.")
            return
        out = os.path.abspath(out)
        self.status.config(text="처리 중…")
        self.update_idletasks()
        try:
            n, skipped = export_docx_paths_to_xlsx(self.paths, out)
        except ValueError as e:
            self.status.config(text="")
            messagebox.showerror("실패", str(e))
            return
        except ImportError as e:
            self.status.config(text="")
            messagebox.showerror("패키지", str(e))
            return
        except OSError as e:
            self.status.config(text="")
            messagebox.showerror("파일 오류", str(e))
            return

        msg = f"{n}행을 저장했습니다.\n\n{out}"
        if skipped:
            msg += "\n\n다음 파일은 코멘트가 없어 건너뜀:\n• " + "\n• ".join(skipped)
        self.status.config(text=f"완료 — {n}행 → {os.path.basename(out)}")
        messagebox.showinfo("완료", msg)


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
