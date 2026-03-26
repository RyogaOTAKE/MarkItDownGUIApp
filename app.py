import os
import threading
import queue
import traceback
from dataclasses import dataclass
from typing import List, Optional, Sequence

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

import sv_ttk
from markitdown import MarkItDown

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None


SUPPORTED_EXTS = {".docx", ".pdf", ".pptx", ".xlsx"}
FILE_DIALOG_TYPES = [
    ("対応ファイル", "*.docx *.pdf *.pptx *.xlsx"),
    ("Word", "*.docx"),
    ("PDF", "*.pdf"),
    ("PowerPoint", "*.pptx"),
    ("Excel", "*.xlsx"),
    ("すべて", "*.*"),
]
BaseApp = TkinterDnD.Tk if TkinterDnD is not None else tk.Tk
_FONT = "Yu Gothic UI"


@dataclass
class Job:
    targets: List[str]
    overwrite: bool
    newline: str
    output_dir: Optional[str]


def list_target_files(folder: str, include_subfolders: bool) -> List[str]:
    files: List[str] = []
    if include_subfolders:
        for root, _, names in os.walk(folder):
            for name in names:
                ext = os.path.splitext(name)[1].lower()
                if ext in SUPPORTED_EXTS:
                    files.append(os.path.join(root, name))
    else:
        for name in os.listdir(folder):
            p = os.path.join(folder, name)
            if not os.path.isfile(p):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext in SUPPORTED_EXTS:
                files.append(p)
    files.sort()
    return files


def output_md_path(input_file: str, output_dir: Optional[str]) -> str:
    base_name = os.path.splitext(os.path.basename(input_file))[0] + ".md"
    if output_dir:
        return os.path.join(output_dir, base_name)
    base, _ = os.path.splitext(input_file)
    return base + ".md"


def convert_one(md: MarkItDown, input_file: str) -> str:
    result = md.convert(input_file)
    text = getattr(result, "text_content", None)
    if text is None:
        raise RuntimeError("変換結果が取得できませんでした。")
    return text


class FileSelectionDialog(tk.Toplevel):
    """フォルダ変換時のファイル確認ダイアログ"""

    def __init__(self, parent: tk.Misc, files: List[str]) -> None:
        super().__init__(parent)
        self.title("変換対象ファイルの確認")
        self.geometry("600x440")
        self.minsize(440, 300)
        self.resizable(True, True)
        self.grab_set()
        self.transient(parent)

        self.result: Optional[List[str]] = None
        self._vars: List[tuple[str, tk.BooleanVar]] = [
            (f, tk.BooleanVar(value=True)) for f in files
        ]

        self._build_ui()
        self.wait_window()

    def _build_ui(self) -> None:
        pad = {"padx": 12}

        # Header
        header = ttk.Frame(self)
        header.pack(fill="x", **pad, pady=(12, 6))
        ttk.Label(
            header,
            text=f"変換対象: {len(self._vars)} 件",
            font=(_FONT, 10, "bold"),
        ).pack(side="left")
        ttk.Button(header, text="全解除", command=self._deselect_all).pack(side="right", padx=(4, 0))
        ttk.Button(header, text="全選択", command=self._select_all).pack(side="right")

        ttk.Separator(self, orient="horizontal").pack(fill="x", **pad, pady=(0, 6))

        # Scrollable checkbox list
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, **pad, pady=(0, 6))

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        canvas = tk.Canvas(list_frame, yscrollcommand=scrollbar.set, borderwidth=0, highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=canvas.yview)

        inner = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_resize(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_resize(e):
            canvas.itemconfig(canvas_window, width=e.width)

        inner.bind("<Configure>", _on_inner_resize)
        canvas.bind("<Configure>", _on_canvas_resize)
        canvas.bind(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )

        for path, var in self._vars:
            row = ttk.Frame(inner)
            row.pack(fill="x", pady=1)
            ttk.Checkbutton(row, variable=var, text=path).pack(
                side="left", anchor="w", padx=4
            )

        # Button row
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", **pad, pady=(0, 12))
        ttk.Button(btn_frame, text="変換開始", command=self._ok).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="キャンセル", command=self._cancel).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _select_all(self) -> None:
        for _, var in self._vars:
            var.set(True)

    def _deselect_all(self) -> None:
        for _, var in self._vars:
            var.set(False)

    def _ok(self) -> None:
        self.result = [f for f, var in self._vars if var.get()]
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


class App(BaseApp):
    def __init__(self) -> None:
        super().__init__()
        self.title("MarkItDown GUI")
        self.geometry("860x620")
        self.minsize(800, 560)

        self._q: "queue.Queue[tuple[str, str]]" = queue.Queue()
        self._stop_event = threading.Event()
        self._worker: Optional[threading.Thread] = None

        # Selection state: "none" | "files" | "folder"
        self._selection_mode: str = "none"
        self._selected_files: List[str] = []
        self._selected_folder: str = ""

        self.var_include_sub = tk.BooleanVar(value=True)
        self.var_overwrite = tk.BooleanVar(value=True)
        self.var_newline_crlf = tk.BooleanVar(value=False)
        self.var_same_dir = tk.BooleanVar(value=True)
        self.var_output_dir = tk.StringVar()

        sv_ttk.set_theme("light")
        self._dark_mode = False
        self._setup_fonts()

        self._build_ui()
        self.after(100, self._poll_queue)

    def _setup_fonts(self) -> None:
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
            try:
                tkfont.nametofont(name).configure(family=_FONT, size=10)
            except Exception:
                pass

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 8}

        # ── ヘッダー ──
        header = ttk.Frame(self)
        header.pack(fill="x", padx=12, pady=(12, 0))
        ttk.Label(header, text="MarkItDown GUI", font=(_FONT, 14, "bold")).pack(side="left")
        self.btn_theme = ttk.Button(
            header, text="ダークモード", width=12, command=self._toggle_theme
        )
        self.btn_theme.pack(side="right")
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12, pady=(8, 0))

        # ── 変換対象 ──
        input_group = ttk.LabelFrame(self, text="変換対象")
        input_group.pack(fill="x", **pad)

        self.lbl_selection = ttk.Label(
            input_group,
            text="ファイルまたはフォルダを選択してください",
            foreground="gray",
        )
        self.lbl_selection.pack(fill="x", padx=10, pady=(8, 4))

        self.file_list = tk.Listbox(
            input_group, height=2, selectmode="browse", font=(_FONT, 9), activestyle="none"
        )
        self.file_list.pack(fill="x", padx=10, pady=(0, 4))

        btn_row = ttk.Frame(input_group)
        btn_row.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Button(btn_row, text="ファイルを選択", command=self._choose_files).pack(side="left")
        ttk.Button(btn_row, text="フォルダを選択", command=self._choose_folder).pack(
            side="left", padx=6
        )
        ttk.Button(btn_row, text="クリア", command=self._clear_selection).pack(side="left")
        self.chk_subfolder = ttk.Checkbutton(
            btn_row, text="サブフォルダも含める", variable=self.var_include_sub
        )

        hint = (
            "ファイルまたはフォルダをこのウィンドウへドロップできます"
            if DND_FILES is not None
            else "D&D: tkinterdnd2 が必要です"
        )
        ttk.Label(input_group, text=hint, foreground="gray", font=(_FONT, 8)).pack(
            anchor="w", padx=10, pady=(0, 8)
        )

        # ── オプション ──
        opt_group = ttk.LabelFrame(self, text="オプション")
        opt_group.pack(fill="x", **pad)

        opt_row1 = ttk.Frame(opt_group)
        opt_row1.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Checkbutton(opt_row1, text="同名.mdがあれば上書き", variable=self.var_overwrite).pack(
            side="left"
        )
        ttk.Checkbutton(
            opt_row1, text="改行をCRLFにする", variable=self.var_newline_crlf
        ).pack(side="left", padx=14)

        opt_row2 = ttk.Frame(opt_group)
        opt_row2.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Checkbutton(
            opt_row2,
            text="入力と同じ場所に出力",
            variable=self.var_same_dir,
            command=self._on_same_dir_toggle,
        ).pack(side="left")
        self.entry_output_dir = ttk.Entry(
            opt_row2, textvariable=self.var_output_dir, state="disabled"
        )
        self.entry_output_dir.pack(side="left", fill="x", expand=True, padx=(10, 6))
        self.btn_output_dir = ttk.Button(
            opt_row2, text="出力先選択", command=self._choose_output_dir, state="disabled"
        )
        self.btn_output_dir.pack(side="left")

        # ── 変換ボタン + プログレス ──
        mid = ttk.Frame(self)
        mid.pack(fill="x", **pad)
        self.btn_convert = ttk.Button(mid, text="変換", command=self._start_conversion, width=10)
        self.btn_convert.pack(side="left")
        self.progress = ttk.Progressbar(mid, orient="horizontal", mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.btn_cancel = ttk.Button(
            mid, text="キャンセル", command=self._cancel, state="disabled"
        )
        self.btn_cancel.pack(side="left")

        # ── ログ ──
        log_group = ttk.LabelFrame(self, text="ログ")
        log_group.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        log_toolbar = ttk.Frame(log_group)
        log_toolbar.pack(fill="x", padx=10, pady=(6, 0))
        ttk.Button(log_toolbar, text="クリア", command=self._clear_log).pack(side="right")

        self.txt = tk.Text(log_group, height=10, wrap="none", font=(_FONT, 9))
        self.txt.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.txt.tag_config("error", foreground="#c0392b")
        self.txt.tag_config("success", foreground="#27ae60")
        self.txt.tag_config("summary", foreground="#2980b9")
        self.txt.tag_config("warn", foreground="#e67e22")

        self._refresh_selection_ui()
        self._log("対象拡張子: .docx .pdf .pptx .xlsx")
        if DND_FILES is not None:
            self._register_drop_targets()

    # ── Theme ──────────────────────────────────────────

    def _toggle_theme(self) -> None:
        self._dark_mode = not self._dark_mode
        sv_ttk.set_theme("dark" if self._dark_mode else "light")
        self.btn_theme.config(text="ライトモード" if self._dark_mode else "ダークモード")

    def _on_same_dir_toggle(self) -> None:
        state = "disabled" if self.var_same_dir.get() else "normal"
        self.entry_output_dir.config(state=state)
        self.btn_output_dir.config(state=state)

    # ── Selection ──────────────────────────────────────

    def _clear_selection(self) -> None:
        self._selection_mode = "none"
        self._selected_files = []
        self._selected_folder = ""
        self._refresh_selection_ui()

    def _set_files(self, paths: List[str], source: str) -> None:
        normalized = self._normalize_paths(paths)
        if not normalized:
            return
        self._selection_mode = "files"
        self._selected_files = normalized
        self._selected_folder = ""
        self._refresh_selection_ui()
        if len(normalized) == 1:
            self._log(f"{source}: {normalized[0]}")
        else:
            self._log(f"{source}: {len(normalized)} 件のファイルを選択")

    def _set_folder(self, path: str, source: str) -> None:
        self._selection_mode = "folder"
        self._selected_files = []
        self._selected_folder = os.path.normpath(path)
        self._refresh_selection_ui()
        self._log(f"{source}: {self._selected_folder}")

    def _refresh_selection_ui(self) -> None:
        self.file_list.delete(0, "end")
        self.chk_subfolder.pack_forget()

        if self._selection_mode == "none":
            self.lbl_selection.config(
                text="ファイルまたはフォルダを選択してください", foreground="gray"
            )
            self.file_list.config(height=1)

        elif self._selection_mode == "files":
            count = len(self._selected_files)
            self.lbl_selection.config(text=f"ファイル {count} 件を選択中", foreground="")
            for p in self._selected_files:
                self.file_list.insert("end", p)
            self.file_list.config(height=min(5, max(2, count)))

        else:  # folder
            self.lbl_selection.config(
                text=f"フォルダ: {self._selected_folder}", foreground=""
            )
            self.file_list.insert("end", self._selected_folder)
            self.file_list.config(height=1)
            self.chk_subfolder.pack(side="right")

    # ── Input handlers ──────────────────────────────────

    def _choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="変換するファイルを選択", filetypes=FILE_DIALOG_TYPES
        )
        if not paths:
            return
        supported = self._filter_supported(list(paths))
        if not supported:
            messagebox.showwarning("対象外", "対応ファイルが含まれていませんでした。")
            return
        self._set_files(supported, "ファイル選択")

    def _choose_folder(self) -> None:
        path = filedialog.askdirectory(title="変換するフォルダを選択")
        if path:
            self._set_folder(path, "フォルダ選択")

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.var_output_dir.set(path)

    def _get_output_dir(self) -> Optional[str]:
        if self.var_same_dir.get():
            return None
        d = self.var_output_dir.get().strip()
        return d if d else None

    # ── D&D ─────────────────────────────────────────────

    def _register_drop_targets(self) -> None:
        widgets = [self, self.file_list, self.lbl_selection]
        for w in widgets:
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", self._handle_drop)

    def _handle_drop(self, event) -> None:
        items = self._normalize_paths(self.tk.splitlist(event.data))
        if not items:
            return

        files = [p for p in items if os.path.isfile(p)]
        folders = [p for p in items if os.path.isdir(p)]
        missing = [p for p in items if not os.path.exists(p)]

        if missing:
            self._log(f"存在しない項目は無視しました: {', '.join(missing)}", "warn")

        if files and folders:
            messagebox.showwarning("未対応", "ファイルとフォルダの同時ドロップには対応していません。")
            return

        if folders:
            if len(folders) > 1:
                self._log("複数フォルダがドロップされたため、先頭の 1 件を採用しました。", "warn")
            self._set_folder(folders[0], "ドロップ")
            return

        supported = self._filter_supported(files)
        unsupported = [p for p in files if p not in supported]
        if unsupported:
            self._log(
                f"非対応ファイルを除外: {', '.join(os.path.basename(p) for p in unsupported)}",
                "warn",
            )
        if not supported:
            messagebox.showwarning("対象外", "対応ファイルが見つかりませんでした。")
            return
        self._set_files(supported, "ドロップ")

    # ── Conversion ──────────────────────────────────────

    def _start_conversion(self) -> None:
        if self._selection_mode == "none":
            messagebox.showwarning("未選択", "ファイルまたはフォルダを選択してください。")
            return

        output_dir = self._get_output_dir()
        if output_dir and not os.path.isdir(output_dir):
            messagebox.showerror("エラー", "出力先フォルダが存在しません。")
            return

        if self._selection_mode == "folder":
            files = list_target_files(self._selected_folder, self.var_include_sub.get())
            if not files:
                messagebox.showinfo("対象なし", "対応ファイルが見つかりませんでした。")
                return
            dlg = FileSelectionDialog(self, files)
            targets = dlg.result
            if targets is None:
                return
            if not targets:
                messagebox.showwarning("未選択", "変換するファイルがありません。")
                return
        else:
            targets = list(self._selected_files)

        self._start_job(
            Job(
                targets=targets,
                overwrite=bool(self.var_overwrite.get()),
                newline="\r\n" if self.var_newline_crlf.get() else "\n",
                output_dir=output_dir,
            )
        )

    def _start_job(self, job: Job) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("実行中", "変換が実行中です。")
            return
        self._stop_event.clear()
        self.btn_cancel.config(state="normal")
        self.btn_convert.config(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = 1
        self._worker = threading.Thread(target=self._run_job, args=(job,), daemon=True)
        self._worker.start()

    def _cancel(self) -> None:
        self._stop_event.set()
        self._log("キャンセル要求を受け付けました。")

    def _run_job(self, job: Job) -> None:
        try:
            md = MarkItDown()
            files = job.targets

            if not files:
                self._q.put(("info", "対象ファイルが見つかりませんでした。"))
                self._q.put(("done", "0,0,0"))
                return

            self._q.put(("set_max", str(len(files))))
            self._q.put(("info", f"変換対象: {len(files)} 件"))

            done_count = 0
            success_count = 0
            skip_count = 0
            error_count = 0

            for f in files:
                if self._stop_event.is_set():
                    self._q.put(("info", "キャンセルしました。"))
                    break

                out_md = output_md_path(f, job.output_dir)
                if (not job.overwrite) and os.path.exists(out_md):
                    self._q.put(("info", f"スキップ(既存): {out_md}"))
                    skip_count += 1
                    done_count += 1
                    self._q.put(("progress", str(done_count)))
                    continue

                try:
                    self._q.put(("info", f"変換中: {f}"))
                    text = convert_one(md, f)
                    with open(out_md, "w", encoding="utf-8", newline=job.newline) as w:
                        w.write(text)
                    self._q.put(("success", f"出力: {out_md}"))
                    success_count += 1
                except Exception as e:
                    self._q.put(("error", f"失敗: {f}\n  {e}"))
                    self._q.put(("error", traceback.format_exc()))
                    error_count += 1

                done_count += 1
                self._q.put(("progress", str(done_count)))

            self._q.put(("done", f"{success_count},{skip_count},{error_count}"))
        except Exception:
            self._q.put(("error", "致命的エラーが発生しました。"))
            self._q.put(("error", traceback.format_exc()))
            self._q.put(("done", "0,0,1"))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "info":
                    self._log(payload)
                elif kind == "error":
                    self._log(payload, "error")
                elif kind == "success":
                    self._log(payload, "success")
                elif kind == "set_max":
                    try:
                        m = int(payload)
                    except Exception:
                        m = 1
                    self.progress["maximum"] = max(1, m)
                    self.progress["value"] = 0
                elif kind == "progress":
                    try:
                        v = int(payload)
                    except Exception:
                        v = 0
                    self.progress["value"] = v
                elif kind == "done":
                    self.btn_cancel.config(state="disabled")
                    self.btn_convert.config(state="normal")
                    try:
                        s, sk, e = (int(x) for x in payload.split(","))
                        summary = f"── 完了: 成功 {s} 件 / スキップ {sk} 件 / 失敗 {e} 件 ──"
                    except Exception:
                        summary = "完了しました。"
                    self._log(summary, "summary")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _log(self, msg: str, tag: str = "") -> None:
        if tag:
            self.txt.insert("end", msg + "\n", tag)
        else:
            self.txt.insert("end", msg + "\n")
        self.txt.see("end")

    def _clear_log(self) -> None:
        self.txt.delete("1.0", "end")

    # ── Helpers ─────────────────────────────────────────

    def _filter_supported(self, paths: Sequence[str]) -> List[str]:
        return [p for p in paths if os.path.splitext(p)[1].lower() in SUPPORTED_EXTS]

    def _normalize_paths(self, paths: Sequence[str]) -> List[str]:
        seen: set = set()
        result: List[str] = []
        for p in paths:
            n = os.path.normpath(p)
            if n not in seen:
                seen.add(n)
                result.append(n)
        return result


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
