"""
三方对比报告生成工具 — 入口脚本

将供应商质量证明书与内部机械性能检测报告进行三方对比，
输出包含 7 列的对比 Excel：
    检测项 / 标准值 / 供应商数值 / 内部检测值 / 是否合格 / 供应商不合格原因 / 内部不合格原因

用法：
    conda activate paddleocr310

    # CLI 模式
    python compare_extract.py -s 供应商.pdf -n 内部检测.pdf [-o 输出目录]

    # GUI 模式（无参数启动）
    python compare_extract.py

作为模块调用：
    from mechanical.compare_pipeline import run_comparison
    xlsx = run_comparison("供应商.pdf", "内部检测.pdf", "output/")
"""

import argparse
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from mechanical.compare_pipeline import run_comparison


def _parse_args():
    try:
        from ocr_config import load_config
        default_output = load_config()["output"]["output_dir"]
    except Exception:
        default_output = "output2"

    parser = argparse.ArgumentParser(
        description="三方对比报告 — 供应商检测值 vs 标准值 vs 内部检测值",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("-s", "--supplier", required=True,
                        help="供应商质量证明书 PDF / 图片路径")
    parser.add_argument("-n", "--internal", required=True,
                        help="内部机械性能检测报告 PDF / 图片路径")
    parser.add_argument("-o", "--output", default=default_output,
                        help="输出目录")
    return parser.parse_args()


# ─── GUI ──────────────────────────────────────────────────────────────────────

class _QueueWriter:
    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, text: str):
        if text:
            self._q.put(text)

    def flush(self):
        pass


class CompareApp(tk.Tk):
    """三方对比报告生成 GUI 应用。"""

    def __init__(self):
        super().__init__()
        self.title("三方对比报告生成工具")
        self.minsize(760, 580)
        self.resizable(True, True)

        try:
            from ocr_config import load_config
            default_output = load_config()["output"]["output_dir"]
        except Exception:
            default_output = "/Users/project/QMS/PaddleOCR/QMS/output2"

        self._supplier_file = tk.StringVar(
            value="/Users/project/QMS/PaddleOCR/QMS/test_files/锻件质量证明书-3.pdf"
        )
        self._internal_file = tk.StringVar(value="")
        self._output_dir = tk.StringVar(value=default_output)
        self._remove_stamp = tk.BooleanVar(value=True)
        self._log_queue: queue.Queue = queue.Queue()
        self._running = False

        self._build_ui()
        self._poll_log()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 5}

        # ── 供应商文件 ──────────────────────────────────────────
        frm_s = ttk.LabelFrame(self, text="① 供应商文件（质量证明书 PDF）", padding=8)
        frm_s.pack(fill="x", **pad)
        ttk.Entry(frm_s, textvariable=self._supplier_file,
                  font=("Menlo", 12)).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(frm_s, text="选择文件…",
                   command=lambda: self._pick_file(self._supplier_file), width=10).pack(side="left")

        # ── 内部检测文件 ────────────────────────────────────────
        frm_n = ttk.LabelFrame(self, text="② 内部检测文件（机械性能检测报告 PDF）", padding=8)
        frm_n.pack(fill="x", **pad)
        ttk.Entry(frm_n, textvariable=self._internal_file,
                  font=("Menlo", 12)).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(frm_n, text="选择文件…",
                   command=lambda: self._pick_file(self._internal_file), width=10).pack(side="left")

        # ── 输出目录 ────────────────────────────────────────────
        frm_out = ttk.LabelFrame(self, text="输出目录", padding=8)
        frm_out.pack(fill="x", **pad)
        ttk.Entry(frm_out, textvariable=self._output_dir,
                  font=("Menlo", 12)).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(frm_out, text="选择目录…", command=self._pick_output).pack(side="left", padx=(0, 4))
        ttk.Button(frm_out, text="打开目录", command=self._open_output).pack(side="left")

        # ── 按钮栏 ──────────────────────────────────────────────
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=12, pady=(8, 2))
        self._btn_start = ttk.Button(frm_btn, text="▶  开始生成", command=self._start, width=14)
        self._btn_start.pack(side="left", padx=(0, 8))
        self._btn_stop = ttk.Button(frm_btn, text="■  停止",
                                    command=self._stop, state="disabled", width=10)
        self._btn_stop.pack(side="left", padx=(0, 12))
        ttk.Checkbutton(frm_btn, text="自动去除公章",
                        variable=self._remove_stamp).pack(side="left", padx=(0, 12))
        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(frm_btn, textvariable=self._status_var, foreground="gray").pack(side="left")

        # ── 进度条 ──────────────────────────────────────────────
        self._progress = ttk.Progressbar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=12, pady=(2, 4))

        # ── 日志 ────────────────────────────────────────────────
        frm_log = ttk.LabelFrame(self, text="运行日志", padding=8)
        frm_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self._log = scrolledtext.ScrolledText(
            frm_log, state="disabled", font=("Menlo", 11),
            wrap="none", relief="flat", bd=0,
        )
        self._log.pack(fill="both", expand=True)

    def _pick_file(self, var: tk.StringVar):
        f = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[
                ("支持的格式", "*.pdf *.jpg *.jpeg *.png"),
                ("PDF 文件", "*.pdf"),
                ("图片文件", "*.jpg *.jpeg *.png"),
            ],
        )
        if f:
            var.set(f)

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
        supplier = self._supplier_file.get().strip()
        internal = self._internal_file.get().strip()
        output = self._output_dir.get().strip()

        if not supplier:
            messagebox.showwarning("提示", "请选择供应商文件。")
            return
        if not internal:
            messagebox.showwarning("提示", "请选择内部检测文件。")
            return
        if not output:
            messagebox.showwarning("提示", "请选择输出目录。")
            return

        self._running = True
        self._btn_start.config(state="disabled")
        self._btn_stop.config(state="normal")
        self._progress.start(12)
        self._status_var.set("生成中…")
        self._log_clear()
        threading.Thread(target=self._run, daemon=True).start()

    def _stop(self):
        self._running = False
        self._btn_stop.config(state="disabled")
        self._status_var.set("正在等待当前任务完成后停止…")

    def _run(self):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = _QueueWriter(self._log_queue)
        sys.stdout = writer
        sys.stderr = writer
        try:
            xlsx = run_comparison(
                supplier_pdf=self._supplier_file.get().strip(),
                internal_pdf=self._internal_file.get().strip(),
                output_dir=self._output_dir.get().strip(),
                remove_stamp=self._remove_stamp.get(),
            )
            if xlsx:
                print(f"\n✓ 报告已生成：{xlsx}")
            else:
                print(f"\n[错误] 未能生成对比报告，请检查输入文件。")
        except Exception as exc:
            import traceback
            print(f"\n[错误] {exc}")
            print(traceback.format_exc())
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


# ─── 入口 ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if "-s" in sys.argv or "--supplier" in sys.argv:
        args = _parse_args()
        run_comparison(args.supplier, args.internal, args.output)
    else:
        CompareApp().mainloop()
