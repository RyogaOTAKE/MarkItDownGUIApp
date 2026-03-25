import os
import threading
import queue
import traceback
from dataclasses import dataclass
from typing import List, Optional, Sequence

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

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


@dataclass
class Job:
    mode: str  # "files" or "folder"
    targets: List[str]
    include_subfolders: bool
    overwrite: bool
    newline: str  # "\n" or "\r\n"


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


def output_md_path(input_file: str) -> str:
    base, _ = os.path.splitext(input_file)
    return base + ".md"


def convert_one(md: MarkItDown, input_file: str) -> str:
    result = md.convert(input_file)
    text = getattr(result, "text_content", None)
    if text is None:
        raise RuntimeError("変換結果が取得できませんでした。")
    return text


class App(BaseApp):
    def __init__(self) -> None:
        super().__init__()
        self.title("MarkItDown GUI")
        self.geometry("860x520")

        self._q: "queue.Queue[tuple[str, str]]" = queue.Queue()
        self._stop_event = threading.Event()
        self._worker: Optional[threading.Thread] = None

        self.selected_files: List[str] = []
        self.var_file_path = tk.StringVar()
        self.var_folder_path = tk.StringVar()
        self.var_include_sub = tk.BooleanVar(value=True)
        self.var_overwrite = tk.BooleanVar(value=True)
        self.var_newline_crlf = tk.BooleanVar(value=False)

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 8}

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        file_group = ttk.LabelFrame(top, text="ファイル変換")
        file_group.pack(fill="x")

        file_row = ttk.Frame(file_group)
        file_row.pack(fill="x", padx=10, pady=8)

        self.file_entry = ttk.Entry(file_row, textvariable=self.var_file_path, state="readonly")
        self.file_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(file_row, text="選択", command=self._choose_files).pack(side="left", padx=6)
        ttk.Button(file_row, text="変換", command=self._start_files).pack(side="left")

        self.file_hint = ttk.Label(
            file_group,
            text="単一ファイル / 複数ファイルを選択できます。ウィンドウへのドラッグ & ドロップにも対応します。",
        )
        self.file_hint.pack(fill="x", padx=10)

        self.file_list = tk.Listbox(file_group, height=4)
        self.file_list.pack(fill="x", padx=10, pady=(6, 10))

        folder_group = ttk.LabelFrame(top, text="フォルダ一括変換")
        folder_group.pack(fill="x", pady=(10, 0))

        folder_row = ttk.Frame(folder_group)
        folder_row.pack(fill="x", padx=10, pady=8)

        self.folder_entry = ttk.Entry(folder_row, textvariable=self.var_folder_path)
        self.folder_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(folder_row, text="選択", command=self._choose_folder).pack(side="left", padx=6)
        ttk.Button(folder_row, text="一括変換", command=self._start_batch).pack(side="left")

        self.folder_hint = ttk.Label(
            folder_group,
            text="フォルダをウィンドウへドラッグ & ドロップすると、この欄に反映されます。",
        )
        self.folder_hint.pack(fill="x", padx=10, pady=(0, 8))

        options_row = ttk.Frame(folder_group)
        options_row.pack(fill="x", padx=10, pady=(0, 8))

        ttk.Checkbutton(options_row, text="サブフォルダも含める", variable=self.var_include_sub).pack(side="left")

        drop_group = ttk.LabelFrame(self, text="ドラッグ & ドロップ")
        drop_group.pack(fill="x", **pad)

        self.drop_label = ttk.Label(
            drop_group,
            text="ファイルまたはフォルダを、このウィンドウへそのままドロップしてください。",
        )
        self.drop_label.pack(fill="x", padx=10, pady=8)

        opt_group = ttk.LabelFrame(self, text="共通オプション")
        opt_group.pack(fill="x", **pad)

        opt_row = ttk.Frame(opt_group)
        opt_row.pack(fill="x", padx=10, pady=8)

        ttk.Checkbutton(opt_row, text="同名.mdがあれば上書き", variable=self.var_overwrite).pack(side="left")
        ttk.Checkbutton(opt_row, text="改行をCRLFにする", variable=self.var_newline_crlf).pack(side="left", padx=14)

        mid = ttk.Frame(self)
        mid.pack(fill="x", **pad)

        self.progress = ttk.Progressbar(mid, orient="horizontal", mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)

        self.btn_cancel = ttk.Button(mid, text="キャンセル", command=self._cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=8)

        log_group = ttk.LabelFrame(self, text="ログ")
        log_group.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.txt = tk.Text(log_group, height=14, wrap="none")
        self.txt.pack(fill="both", expand=True, padx=10, pady=10)

        self._update_file_selection([])
        self._log("対象拡張子: .docx .pdf .pptx .xlsx")
        if DND_FILES is None:
            self._log("ドラッグ & ドロップ機能は無効です。tkinterdnd2 をインストールすると有効になります。")
            self.drop_label.configure(text="ドラッグ & ドロップを使うには、tkinterdnd2 のインストールが必要です。")
        else:
            self._register_drop_targets()
            self._log("ファイル / フォルダのドラッグ & ドロップに対応しています。")

    def _register_drop_targets(self) -> None:
        widgets = [self, self.file_entry, self.file_list, self.folder_entry, self.drop_label]
        for widget in widgets:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._handle_drop)

    def _choose_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="変換するファイルを選択",
            filetypes=FILE_DIALOG_TYPES,
        )
        if paths:
            self._set_selected_files(list(paths), source="選択")

    def _choose_folder(self) -> None:
        path = filedialog.askdirectory(title="変換するフォルダを選択")
        if path:
            self.var_folder_path.set(path)
            self._log(f"フォルダを選択しました: {path}")

    def _start_files(self) -> None:
        paths = list(self.selected_files)
        if not paths:
            messagebox.showwarning("入力不足", "ファイルを選択してください。")
            return
        missing_files = [path for path in paths if not os.path.isfile(path)]
        if missing_files:
            messagebox.showerror("エラー", "存在しないファイルが含まれています。")
            return
        self._start_job(Job(
            mode="files",
            targets=paths,
            include_subfolders=False,
            overwrite=bool(self.var_overwrite.get()),
            newline="\r\n" if self.var_newline_crlf.get() else "\n",
        ))

    def _start_batch(self) -> None:
        path = self.var_folder_path.get().strip()
        if not path:
            messagebox.showwarning("入力不足", "フォルダを選択してください。")
            return
        if not os.path.isdir(path):
            messagebox.showerror("エラー", "フォルダが存在しません。")
            return
        self._start_job(Job(
            mode="folder",
            targets=[path],
            include_subfolders=bool(self.var_include_sub.get()),
            overwrite=bool(self.var_overwrite.get()),
            newline="\r\n" if self.var_newline_crlf.get() else "\n",
        ))

    def _start_job(self, job: Job) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("実行中", "変換が実行中です。完了またはキャンセルしてください。")
            return

        self._stop_event.clear()
        self.btn_cancel.config(state="normal")
        self.progress["value"] = 0
        self.progress["maximum"] = 1

        self._worker = threading.Thread(target=self._run_job, args=(job,), daemon=True)
        self._worker.start()

    def _cancel(self) -> None:
        self._stop_event.set()
        self._log("キャンセル要求を受け付けました。現在処理中のファイル完了後に停止します。")

    def _run_job(self, job: Job) -> None:
        try:
            md = MarkItDown()

            if job.mode == "folder":
                files = list_target_files(job.targets[0], job.include_subfolders)
            else:
                files = list(job.targets)

            if not files:
                self._q.put(("info", "対象ファイルが見つかりませんでした。"))
                self._q.put(("done", ""))
                return

            self._q.put(("set_max", str(len(files))))
            self._q.put(("info", f"変換対象: {len(files)} 件"))

            done_count = 0
            for f in files:
                if self._stop_event.is_set():
                    self._q.put(("info", "キャンセルしました。"))
                    break

                out_md = output_md_path(f)
                if (not job.overwrite) and os.path.exists(out_md):
                    self._q.put(("info", f"スキップ(既存): {out_md}"))
                    done_count += 1
                    self._q.put(("progress", str(done_count)))
                    continue

                try:
                    self._q.put(("info", f"変換中: {f}"))
                    text = convert_one(md, f)

                    with open(out_md, "w", encoding="utf-8", newline=job.newline) as w:
                        w.write(text)

                    self._q.put(("info", f"出力: {out_md}"))
                except Exception as e:
                    self._q.put(("error", f"失敗: {f}\n  {e}"))
                    self._q.put(("error", traceback.format_exc()))

                done_count += 1
                self._q.put(("progress", str(done_count)))

            self._q.put(("done", ""))
        except Exception:
            self._q.put(("error", "致命的エラーが発生しました。"))
            self._q.put(("error", traceback.format_exc()))
            self._q.put(("done", ""))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._q.get_nowait()
                if kind == "info":
                    self._log(payload)
                elif kind == "error":
                    self._log(payload)
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
                    self._log("完了しました。")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _log(self, msg: str) -> None:
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")

    def _handle_drop(self, event) -> None:
        dropped_items = self._normalize_paths(self.tk.splitlist(event.data))
        if not dropped_items:
            return

        dropped_files = [path for path in dropped_items if os.path.isfile(path)]
        dropped_folders = [path for path in dropped_items if os.path.isdir(path)]
        missing_items = [path for path in dropped_items if not os.path.exists(path)]

        if missing_items:
            self._log(f"存在しない項目は無視しました: {', '.join(missing_items)}")

        if dropped_files and dropped_folders:
            messagebox.showwarning("未対応", "ファイルとフォルダの同時ドロップには対応していません。")
            self._log("ファイルとフォルダの同時ドロップは受け付けませんでした。")
            return

        if dropped_folders:
            selected_folder = dropped_folders[0]
            self.var_folder_path.set(selected_folder)
            if len(dropped_folders) > 1:
                self._log("複数フォルダがドロップされたため、先頭の 1 件だけを採用しました。")
            self._log(f"フォルダをドロップしました: {selected_folder}")
            return

        supported_files = self._filter_supported_files(dropped_files)
        unsupported_files = [path for path in dropped_files if path not in supported_files]

        if unsupported_files:
            self._log(f"非対応ファイルは除外しました: {', '.join(unsupported_files)}")

        if not supported_files:
            messagebox.showwarning("対象外", "対応ファイルが見つかりませんでした。")
            self._log("対応ファイルが含まれていないため、選択状態は更新しませんでした。")
            return

        self._set_selected_files(supported_files, source="ドロップ")

    def _set_selected_files(self, paths: Sequence[str], source: str) -> None:
        normalized_paths = self._normalize_paths(paths)
        self._update_file_selection(normalized_paths)

        if not normalized_paths:
            return
        if len(normalized_paths) == 1:
            self._log(f"{source}でファイルを 1 件選択しました: {normalized_paths[0]}")
            return
        self._log(f"{source}でファイルを {len(normalized_paths)} 件選択しました。")

    def _update_file_selection(self, paths: Sequence[str]) -> None:
        self.selected_files = list(paths)
        self.file_list.delete(0, "end")

        for path in self.selected_files:
            self.file_list.insert("end", path)

        visible_rows = min(6, max(1, len(self.selected_files)))
        self.file_list.configure(height=visible_rows)

        if not self.selected_files:
            self.var_file_path.set("ファイル未選択")
        elif len(self.selected_files) == 1:
            self.var_file_path.set(self.selected_files[0])
        else:
            self.var_file_path.set(f"{len(self.selected_files)} 件のファイルを選択中")

    def _filter_supported_files(self, paths: Sequence[str]) -> List[str]:
        supported_files: List[str] = []
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            if ext in SUPPORTED_EXTS:
                supported_files.append(path)
        return supported_files

    def _normalize_paths(self, paths: Sequence[str]) -> List[str]:
        normalized_paths: List[str] = []
        seen_paths = set()

        for path in paths:
            normalized_path = os.path.normpath(path)
            if normalized_path in seen_paths:
                continue
            seen_paths.add(normalized_path)
            normalized_paths.append(normalized_path)

        return normalized_paths


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
