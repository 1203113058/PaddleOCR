"""
PDF OCR 工具 — 基础配置页面

启动方式：
    conda activate paddleocr
    python config_ui.py

浏览器将自动打开 http://localhost:7860，在页面上修改参数后点击「保存配置」即可。
修改结果持久化到 config.toml，下次运行 pdf2excel.py 时自动读取。

依赖：
    pip install gradio
"""

from pathlib import Path

try:
    import gradio as gr
except ImportError:
    raise SystemExit(
        "\n[错误] 未安装 gradio，请先执行：pip install gradio\n"
    )

from ocr_config import (
    CONFIG_PATH,
    DEFAULT_CONFIG,
    load_config,
    save_config,
)


# ─── 工具函数 ────────────────────────────────────────────────────────────────

def _build_cfg(output_dir: str, threshold: float, row_tol: int) -> dict:
    return {
        "output": {"output_dir": output_dir.strip()},
        "ocr":    {"threshold": round(float(threshold), 4), "row_tol": int(row_tol)},
    }


def do_save(output_dir: str, threshold: float, row_tol: int) -> str:
    if not output_dir.strip():
        return "⚠️ 输出目录不能为空，请重新输入。"
    cfg = _build_cfg(output_dir, threshold, row_tol)
    path = save_config(cfg)
    return f"✅ 配置已保存到：{path}"


def do_reset() -> tuple:
    d = DEFAULT_CONFIG
    return (
        d["output"]["output_dir"],
        d["ocr"]["threshold"],
        d["ocr"]["row_tol"],
        f"✅ 已恢复内置默认值（未写入文件）。点击「保存配置」后才会持久化。",
    )


def do_open_folder(output_dir: str) -> str:
    import subprocess, sys
    path = Path(output_dir.strip())
    if not path.is_absolute():
        path = Path(__file__).parent / path
    path.mkdir(parents=True, exist_ok=True)
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    elif sys.platform == "win32":
        subprocess.Popen(["explorer", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])
    return f"已在文件管理器中打开：{path}"


# ─── 界面构建 ────────────────────────────────────────────────────────────────

def build_ui() -> gr.Blocks:
    cfg = load_config()

    with gr.Blocks(title="PDF OCR 配置", theme=gr.themes.Soft()) as app:
        gr.Markdown(
            "## PDF OCR 工具 — 基础配置\n"
            f"配置文件路径：`{CONFIG_PATH}`\n\n"
            "修改下方参数后点击 **保存配置** 即可生效，下次运行 `pdf2excel.py` 时自动读取。"
        )

        with gr.Group():
            gr.Markdown("### 输出设置")
            with gr.Row():
                output_dir = gr.Textbox(
                    label="默认输出目录",
                    value=cfg["output"]["output_dir"],
                    placeholder="支持绝对路径或相对路径，例如 /Users/me/ocr_output",
                    scale=4,
                )
                open_btn = gr.Button("📂 打开目录", scale=1, variant="secondary")

        with gr.Group():
            gr.Markdown("### OCR 参数")
            with gr.Row():
                threshold = gr.Slider(
                    label="低置信度阈值",
                    minimum=0.80,
                    maximum=1.00,
                    step=0.001,
                    value=cfg["ocr"]["threshold"],
                    info="低于此值的文本将在图片和 Excel 中标红",
                )
                row_tol = gr.Slider(
                    label="行分组容忍度（px）",
                    minimum=5,
                    maximum=50,
                    step=1,
                    value=cfg["ocr"]["row_tol"],
                    info="Y 轴差值在此范围内视为同一行，值越大行合并越宽松",
                )

        status = gr.Textbox(label="状态", interactive=False, lines=1)

        with gr.Row():
            save_btn  = gr.Button("💾 保存配置", variant="primary")
            reset_btn = gr.Button("↩️ 恢复默认值", variant="secondary")

        # 事件绑定
        save_btn.click(
            fn=do_save,
            inputs=[output_dir, threshold, row_tol],
            outputs=status,
        )
        reset_btn.click(
            fn=do_reset,
            inputs=None,
            outputs=[output_dir, threshold, row_tol, status],
        )
        open_btn.click(
            fn=do_open_folder,
            inputs=output_dir,
            outputs=status,
        )

        gr.Markdown(
            "---\n"
            "**命令行优先级说明**：命令行显式传入的参数（`-o`, `--threshold`, `--row-tol`）"
            "始终覆盖此处配置文件中的默认值。"
        )

    return app


if __name__ == "__main__":
    ui = build_ui()
    ui.launch(inbrowser=True)
