"""
机械性能提取 — GUI 界面（Tkinter）
"""

import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from .pipeline import extract_mechanical


class _QueueWriter:
    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, text: str):
        if text:
            self._q.put(text)

    def flush(self):
        pass


class MechApp(tk.Tk):
    """机械性能专项提取 GUI 应用。"""

    def __init__(self):
        super().__init__()
        self.title("综合报告 — 机械性能专项提取")
        self.minsize(720, 560)
        self.resizable(True, True)

        self._output_dir = tk.StringVar(value="/Users/project/QMS/PaddleOCR/QMS/output2")
        self._remove_stamp = tk.BooleanVar(value=True)
        self._pdf_files: list[str] = [
            "/Users/project/QMS/PaddleOCR/QMS/test_files/锻件质量证明书-3.pdf"
        ]
        self._log_queue: queue.Queue = queue.Queue()
        self._running = False

        self._build_ui()
        self._poll_log()
        self._refresh_files_text()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        frm_in = ttk.LabelFrame(self, text="输入文件（综合检测报告 PDF）", padding=8)
        frm_in.pack(fill="x", **pad)

        self._files_text = tk.Text(
            frm_in, height=4, state="disabled", wrap="none",
            relief="sunken", bd=1, font=("Menlo", 12),
        )
        self._files_text.pack(side="left", fill="both", expand=True, padx=(0, 8))

        frm_in_btns = ttk.Frame(frm_in)
        frm_in_btns.pack(side="right", fill="y")
        ttk.Button(frm_in_btns, text="选择文件…", command=self._pick_files, width=10).pack(fill="x", pady=(0, 4))
        ttk.Button(frm_in_btns, text="清  空", command=self._clear_files, width=10).pack(fill="x")

        frm_out = ttk.LabelFrame(self, text="输出目录", padding=8)
        frm_out.pack(fill="x", **pad)
        ttk.Entry(frm_out, textvariable=self._output_dir, font=("Menlo", 12)).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(frm_out, text="选择目录…", command=self._pick_output).pack(side="left", padx=(0, 4))
        ttk.Button(frm_out, text="打开目录", command=self._open_output).pack(side="left")

        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=12, pady=(8, 2))
        self._btn_start = ttk.Button(frm_btn, text="▶  开始提取", command=self._start, width=14)
        self._btn_start.pack(side="left", padx=(0, 8))
        self._btn_stop = ttk.Button(frm_btn, text="■  停止", command=self._stop, state="disabled", width=10)
        self._btn_stop.pack(side="left", padx=(0, 12))
        ttk.Checkbutton(frm_btn, text="自动去除公章", variable=self._remove_stamp).pack(side="left", padx=(0, 12))
        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(frm_btn, textvariable=self._status_var, foreground="gray").pack(side="left")

        self._progress = ttk.Progressbar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=12, pady=(2, 4))

        frm_log = ttk.LabelFrame(self, text="运行日志", padding=8)
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self._log = scrolledtext.ScrolledText(
            frm_log, state="disabled", font=("Menlo", 11),
            wrap="none", relief="flat", bd=0,
        )
        self._log.pack(fill="both", expand=True)

    def _pick_files(self):
        files = filedialog.askopenfilenames(
            title="选择综合检测报告",
            filetypes=[
                ("支持的格式", "*.pdf *.jpg *.jpeg *.png"),
                ("PDF 文件", "*.pdf"),
                ("图片文件", "*.jpg *.jpeg *.png"),
            ],
        )
        if files:
            self._pdf_files = list(files)
            self._refresh_files_text()

    def _clear_files(self):
        self._pdf_files = []
        self._refresh_files_text()

    def _refresh_files_text(self):
        self._files_text.config(state="normal")
        self._files_text.delete("1.0", "end")
        for f in self._pdf_files:
            self._files_text.insert("end", f + "\n")
        self._files_text.config(state="disabled")

    def _pick_output(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self._output_dir.set(d)

    def _open_output(self):
        raw = self._output_dir.get().strip()
        path = Path(raw) if Path(raw).is_absolute() else Path(__file__).parent / raw
        path.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["open", str(path)])

    def _start(self):
        if not self._pdf_files:
            messagebox.showwarning("提示", "请先选择至少一个文件。")
            return
        if not self._output_dir.get().strip():
            messagebox.showwarning("提示", "请选择输出目录。")
            return
        self._running = True
        self._btn_start.config(state="disabled")
        self._btn_stop.config(state="normal")
        self._progress.start(12)
        self._status_var.set("提取中…")
        self._log_clear()
        threading.Thread(target=self._run, daemon=True).start()

    def _stop(self):
        self._running = False
        self._btn_stop.config(state="disabled")
        self._status_var.set("正在等待当前文件完成后停止…")

    def _run(self):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = _QueueWriter(self._log_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            total = len(self._pdf_files)
            completed = 0
            for idx, pdf_path in enumerate(self._pdf_files, 1):
                if not self._running:
                    print("\n[停止] 已中止。")
                    break
                print(f"\n{'═'*55}")
                print(f"[{idx}/{total}]  {Path(pdf_path).name}")
                print(f"{'═'*55}")
                try:
                    extract_mechanical(
                        pdf_path,
                        self._output_dir.get().strip(),
                        remove_stamp=self._remove_stamp.get(),
                    )
                    completed += 1
                except Exception as exc:
                    print(f"\n[错误] 处理失败：{exc}")
            if self._running:
                print(f"\n{'═'*55}")
                print(f"✓ 全部完成：{completed}/{total} 个文件处理成功。")
                print(f"{'═'*55}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.after(0, self._on_done)

    def _on_done(self):
        self._running = False
        self._progress.stop()
        self._btn_start.config(state="normal")
        self._btn_stop.config(state="disabled")
        self._status_var.set("完成")

    def _log_clear(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _poll_log(self):
        try:
            while True:
                text = self._log_queue.get_nowait()
                self._log.config(state="normal")
                self._log.insert("end", text)
                self._log.see("end")
                self._log.config(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._poll_log)
