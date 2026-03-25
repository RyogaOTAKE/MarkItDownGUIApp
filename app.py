import os
import threading
import queue
import traceback
from dataclasses import dataclass
from typing import List, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import sv_ttk
from markitdown import MarkItDown


SUPPORTED_EXTS = {".docx", ".pdf", ".pptx", ".xlsx"}


@dataclass
class Job:
    mode: str  # "file" or "folder"
    input_path: str
    include_subfolders: bool
    overwrite: bool
    newline: str  # "\n" or "\r\n"
    output_dir: Optional[str]  # None = same as input


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


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("MarkItDown GUI")
        self.geometry("860x620")
        self.minsize(800, 560)

        self._q: "queue.Queue[tuple[str, str]]" = queue.Queue()
        self._stop_event = threading.Event()
        self._worker: Optional[threading.Thread] = None

        self.var_file_path = tk.StringVar()
        self.var_folder_path = tk.StringVar()
        self.var_include_sub = tk.BooleanVar(value=True)
        self.var_overwrite = tk.BooleanVar(value=True)
        self.var_newline_crlf = tk.BooleanVar(value=False)
        self.var_same_dir = tk.BooleanVar(value=True)
        self.var_output_dir = tk.StringVar()

        sv_ttk.set_theme("light")
        self._dark_mode = False

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 8}

        # ── ヘッダー（タイトル + テーマトグル）──
        header = ttk.Frame(self)
        header.pack(fill="x", padx=10, pady=(10, 0))

        ttk.Label(header, text="MarkItDown GUI", font=("", 13, "bold")).pack(side="left")
        self.btn_theme = ttk.Button(header, text="ダークモード", width=12, command=self._toggle_theme)
        self.btn_theme.pack(side="right")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=10, pady=(6, 0))

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        # 単一ファイル
        file_group = ttk.LabelFrame(top, text="単一ファイル変換")
        file_group.pack(fill="x")

        file_row = ttk.Frame(file_group)
        file_row.pack(fill="x", padx=10, pady=8)

        ttk.Entry(file_row, textvariable=self.var_file_path).pack(side="left", fill="x", expand=True)
        ttk.Button(file_row, text="選択", command=self._choose_file).pack(side="left", padx=6)
        ttk.Button(file_row, text="変換", command=self._start_single).pack(side="left")

        # フォルダ一括
        folder_group = ttk.LabelFrame(top, text="フォルダ一括変換")
        folder_group.pack(fill="x", pady=(10, 0))

        folder_row = ttk.Frame(folder_group)
        folder_row.pack(fill="x", padx=10, pady=8)

        ttk.Entry(folder_row, textvariable=self.var_folder_path).pack(side="left", fill="x", expand=True)
        ttk.Button(folder_row, text="選択", command=self._choose_folder).pack(side="left", padx=6)
        ttk.Button(folder_row, text="一括変換", command=self._start_batch).pack(side="left")

        options_row = ttk.Frame(folder_group)
        options_row.pack(fill="x", padx=10, pady=(0, 8))

        ttk.Checkbutton(options_row, text="サブフォルダも含める", variable=self.var_include_sub).pack(side="left")

        # 共通オプション
        opt_group = ttk.LabelFrame(self, text="共通オプション")
        opt_group.pack(fill="x", **pad)

        opt_row1 = ttk.Frame(opt_group)
        opt_row1.pack(fill="x", padx=10, pady=(8, 4))

        ttk.Checkbutton(opt_row1, text="同名.mdがあれば上書き", variable=self.var_overwrite).pack(side="left")
        ttk.Checkbutton(opt_row1, text="改行をCRLFにする", variable=self.var_newline_crlf).pack(side="left", padx=14)

        opt_row2 = ttk.Frame(opt_group)
        opt_row2.pack(fill="x", padx=10, pady=(0, 8))

        self.chk_same_dir = ttk.Checkbutton(
            opt_row2,
            text="入力と同じ場所に出力",
            variable=self.var_same_dir,
            command=self._on_same_dir_toggle,
        )
        self.chk_same_dir.pack(side="left")

        self.entry_output_dir = ttk.Entry(opt_row2, textvariable=self.var_output_dir, state="disabled")
        self.entry_output_dir.pack(side="left", fill="x", expand=True, padx=(10, 6))

        self.btn_output_dir = ttk.Button(
            opt_row2, text="出力先選択", command=self._choose_output_dir, state="disabled"
        )
        self.btn_output_dir.pack(side="left")

        # 進捗と操作
        mid = ttk.Frame(self)
        mid.pack(fill="x", **pad)

        self.progress = ttk.Progressbar(mid, orient="horizontal", mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True)

        self.btn_cancel = ttk.Button(mid, text="キャンセル", command=self._cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=8)

        # ログ
        log_group = ttk.LabelFrame(self, text="ログ")
        log_group.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        log_toolbar = ttk.Frame(log_group)
        log_toolbar.pack(fill="x", padx=10, pady=(6, 0))

        ttk.Button(log_toolbar, text="クリア", command=self._clear_log).pack(side="right")

        self.txt = tk.Text(log_group, height=14, wrap="none")
        self.txt.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        self.txt.tag_config("error", foreground="#e05252")
        self.txt.tag_config("success", foreground="#3a9e5f")
        self.txt.tag_config("summary", foreground="#1e88e5")

        self._log("対象拡張子: .docx .pdf .pptx .xlsx")

    def _toggle_theme(self) -> None:
        self._dark_mode = not self._dark_mode
        sv_ttk.set_theme("dark" if self._dark_mode else "light")
        self.btn_theme.config(text="ライトモード" if self._dark_mode else "ダークモード")

    def _on_same_dir_toggle(self) -> None:
        if self.var_same_dir.get():
            self.entry_output_dir.config(state="disabled")
            self.btn_output_dir.config(state="disabled")
        else:
            self.entry_output_dir.config(state="normal")
            self.btn_output_dir.config(state="normal")

    def _choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="変換するファイルを選択",
            filetypes=[
                ("対応ファイル", "*.docx *.pdf *.pptx *.xlsx"),
                ("Word", "*.docx"),
                ("PDF", "*.pdf"),
                ("PowerPoint", "*.pptx"),
                ("Excel", "*.xlsx"),
                ("すべて", "*.*"),
            ],
        )
        if path:
            self.var_file_path.set(path)

    def _choose_folder(self) -> None:
        path = filedialog.askdirectory(title="変換するフォルダを選択")
        if path:
            self.var_folder_path.set(path)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.var_output_dir.set(path)

    def _get_output_dir(self) -> Optional[str]:
        if self.var_same_dir.get():
            return None
        d = self.var_output_dir.get().strip()
        return d if d else None

    def _start_single(self) -> None:
        path = self.var_file_path.get().strip()
        if not path:
            messagebox.showwarning("入力不足", "ファイルを選択してください。")
            return
        if not os.path.isfile(path):
            messagebox.showerror("エラー", "ファイルが存在しません。")
            return
        output_dir = self._get_output_dir()
        if output_dir and not os.path.isdir(output_dir):
            messagebox.showerror("エラー", "出力先フォルダが存在しません。")
            return
        self._start_job(Job(
            mode="file",
            input_path=path,
            include_subfolders=False,
            overwrite=bool(self.var_overwrite.get()),
            newline="\r\n" if self.var_newline_crlf.get() else "\n",
            output_dir=output_dir,
        ))

    def _start_batch(self) -> None:
        path = self.var_folder_path.get().strip()
        if not path:
            messagebox.showwarning("入力不足", "フォルダを選択してください。")
            return
        if not os.path.isdir(path):
            messagebox.showerror("エラー", "フォルダが存在しません。")
            return
        output_dir = self._get_output_dir()
        if output_dir and not os.path.isdir(output_dir):
            messagebox.showerror("エラー", "出力先フォルダが存在しません。")
            return
        self._start_job(Job(
            mode="folder",
            input_path=path,
            include_subfolders=bool(self.var_include_sub.get()),
            overwrite=bool(self.var_overwrite.get()),
            newline="\r\n" if self.var_newline_crlf.get() else "\n",
            output_dir=output_dir,
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

            if job.mode == "file":
                files = [job.input_path]
            else:
                files = list_target_files(job.input_path, job.include_subfolders)

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
                    self._log(payload, tag="error")
                elif kind == "success":
                    self._log(payload, tag="success")
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
                    try:
                        s, sk, e = (int(x) for x in payload.split(","))
                        summary = f"── 完了: 成功 {s} 件 / スキップ {sk} 件 / 失敗 {e} 件 ──"
                    except Exception:
                        summary = "完了しました。"
                    self._log(summary, tag="summary")
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


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
