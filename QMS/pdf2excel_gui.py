"""
PDF 表格识别 GUI 启动器

使用方式：
    conda activate paddleocr310
    python pdf2excel_gui.py

功能：
  · 图形界面选择输入 PDF 文件（支持多文件批量处理）
  · 图形界面选择输出目录
  · 调整 OCR 参数（置信度阈值、行容忍度）
  · 实时日志输出到界面
  · 完成后一键打开输出目录
"""

import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from ocr_config import load_config


# ─── stdout 重定向 ────────────────────────────────────────────────────────────

class _QueueWriter:
    """将 stdout/stderr 写入线程安全队列，供 GUI 主线程消费。"""

    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, text: str):
        if text:
            self._q.put(text)

    def flush(self):
        pass


# ─── 主界面 ───────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 表格识别工具")
        self.minsize(720, 600)
        self.resizable(True, True)

        cfg = load_config()
        cfg_out = cfg["output"]["output_dir"]
        if cfg_out == "output":
            cfg_out = "/Users/project/QMS/PaddleOCR/QMS/output"
        self._output_dir  = tk.StringVar(value=cfg_out)
        self._threshold   = tk.DoubleVar(value=cfg["ocr"]["threshold"])
        self._thr_display = tk.StringVar(value=f'{cfg["ocr"]["threshold"]:.3f}')
        self._row_tol     = tk.IntVar(value=cfg["ocr"]["row_tol"])
        self._pdf_files: list[str] = []
        self._log_queue: queue.Queue = queue.Queue()
        self._running = False

        # 阈值滑块变动时同步显示文本
        self._threshold.trace_add("write", self._on_threshold_change)

        self._build_ui()
        self._poll_log()

    # ─── UI 构建 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── 输入文件区 ──
        frm_in = ttk.LabelFrame(self, text="输入 PDF 文件", padding=8)
        frm_in.pack(fill="x", **pad)

        self._files_text = tk.Text(
            frm_in, height=4, state="disabled",
            wrap="none", relief="sunken", bd=1,
            font=("Menlo", 12),
        )
        self._files_text.pack(side="left", fill="both", expand=True, padx=(0, 8))

        frm_in_btns = ttk.Frame(frm_in)
        frm_in_btns.pack(side="right", fill="y")
        ttk.Button(frm_in_btns, text="选择文件…", command=self._pick_files, width=10).pack(
            fill="x", pady=(0, 4))
        ttk.Button(frm_in_btns, text="清  空", command=self._clear_files, width=10).pack(fill="x")

        # ── 输出目录区 ──
        frm_out = ttk.LabelFrame(self, text="输出目录", padding=8)
        frm_out.pack(fill="x", **pad)

        ttk.Entry(frm_out, textvariable=self._output_dir, font=("Menlo", 12)).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(frm_out, text="选择目录…", command=self._pick_output).pack(side="left", padx=(0, 4))
        ttk.Button(frm_out, text="打开目录", command=self._open_output).pack(side="left")

        # ── OCR 参数区 ──
        frm_cfg = ttk.LabelFrame(self, text="OCR 参数", padding=8)
        frm_cfg.pack(fill="x", **pad)
        frm_cfg.columnconfigure(1, weight=1)

        ttk.Label(frm_cfg, text="低置信度阈值：").grid(row=0, column=0, sticky="w")
        ttk.Scale(
            frm_cfg, from_=0.80, to=1.00, orient="horizontal",
            variable=self._threshold,
        ).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(frm_cfg, textvariable=self._thr_display, width=6).grid(row=0, column=2)
        ttk.Label(frm_cfg, text="低于此值的文本将标红", foreground="gray").grid(
            row=0, column=3, sticky="w", padx=(4, 0))

        ttk.Label(frm_cfg, text="行分组容忍度（px）：").grid(
            row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Spinbox(
            frm_cfg, from_=5, to=50, textvariable=self._row_tol, width=7,
        ).grid(row=1, column=1, sticky="w", padx=8, pady=(8, 0))
        ttk.Label(frm_cfg, text="Y 轴差值在此范围内视为同一行", foreground="gray").grid(
            row=1, column=3, sticky="w", padx=(4, 0), pady=(8, 0))

        # ── 操作按钮与状态 ──
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=12, pady=(4, 2))

        self._btn_start = ttk.Button(frm_btn, text="▶  开始识别", command=self._start, width=14)
        self._btn_start.pack(side="left", padx=(0, 8))

        self._btn_stop = ttk.Button(frm_btn, text="■  停止", command=self._stop,
                                    state="disabled", width=10)
        self._btn_stop.pack(side="left", padx=(0, 12))

        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(frm_btn, textvariable=self._status_var, foreground="gray").pack(side="left")

        # ── 进度条 ──
        self._progress = ttk.Progressbar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=12, pady=(2, 4))

        # ── 日志区 ──
        frm_log = ttk.LabelFrame(self, text="运行日志", padding=8)
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self._log = scrolledtext.ScrolledText(
            frm_log, state="disabled",
            font=("Menlo", 11), wrap="none",
            relief="flat", bd=0,
        )
        self._log.pack(fill="both", expand=True)

    # ─── 参数联动 ─────────────────────────────────────────────────────────────

    def _on_threshold_change(self, *_):
        try:
            self._thr_display.set(f"{self._threshold.get():.3f}")
        except tk.TclError:
            pass

    # ─── 文件 / 目录选择 ──────────────────────────────────────────────────────

    def _pick_files(self):
        files = filedialog.askopenfilenames(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
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

    # ─── 识别控制 ─────────────────────────────────────────────────────────────

    def _start(self):
        if not self._pdf_files:
            messagebox.showwarning("提示", "请先选择至少一个 PDF 文件。")
            return
        if not self._output_dir.get().strip():
            messagebox.showwarning("提示", "请选择输出目录。")
            return

        self._running = True
        self._btn_start.config(state="disabled")
        self._btn_stop.config(state="normal")
        self._progress.start(12)
        self._status_var.set("识别中…")
        self._log_clear()

        threading.Thread(target=self._run_ocr, daemon=True).start()

    def _stop(self):
        self._running = False
        self._btn_stop.config(state="disabled")
        self._status_var.set("正在等待当前页完成后停止…")

    def _run_ocr(self):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = _QueueWriter(self._log_queue)
        sys.stdout = writer
        sys.stderr = writer

        try:
            from pdf2excel import ocr_pdf_to_excel

            total = len(self._pdf_files)
            completed = 0
            for idx, pdf_path in enumerate(self._pdf_files, 1):
                if not self._running:
                    print("\n[停止] 已中止识别。")
                    break
                print(f"\n{'═'*50}")
                print(f"[{idx}/{total}]  {Path(pdf_path).name}")
                print(f"{'═'*50}")
                try:
                    ocr_pdf_to_excel(
                        pdf_path,
                        self._output_dir.get().strip(),
                        round(self._threshold.get(), 4),
                        self._row_tol.get(),
                    )
                    completed += 1
                except Exception as exc:
                    print(f"\n[错误] 处理失败：{exc}")

            if self._running:
                print(f"\n{'═'*50}")
                print(f"✓ 全部完成：{completed}/{total} 个文件处理成功。")
                print(f"{'═'*50}")
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

    # ─── 日志轮询 ─────────────────────────────────────────────────────────────

    def _log_clear(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    def _poll_log(self):
        """每 100 ms 从队列取出日志文本，写入控件（主线程安全）。"""
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


# ─── 入口 ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
