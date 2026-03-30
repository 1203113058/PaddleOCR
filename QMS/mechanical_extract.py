"""
综合检测报告 — 机械性能（力学）表格专项提取

薄入口脚本：所有逻辑已拆分到 mechanical/ 包中。

用法：
    conda activate paddleocr310

    # 命令行模式
    python mechanical_extract.py -i <PDF路径> [-o <输出目录>]

    # GUI 模式（无参数启动）
    python mechanical_extract.py

示例：
    python mechanical_extract.py -i "test_files/锻件质量证明书-3.pdf"

作为模块调用：
    from mechanical import extract_mechanical
    result = extract_mechanical("报告.pdf", "output/")
    print(result["fields"])    # 结构化字段
    print(result["failures"])  # 不合格项
"""

import argparse
import sys

from mechanical import extract_mechanical, MechApp


def _parse_args():
    try:
        from ocr_config import load_config
        cfg = load_config()
        default_output = cfg["output"]["output_dir"]
    except Exception:
        default_output = "output2"

    parser = argparse.ArgumentParser(
        description="综合检测报告 — 机械性能表格专项提取",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "-i", "--input", required=True,
        help="输入文件路径，支持 PDF / JPG / JPEG / PNG",
    )
    parser.add_argument(
        "-o", "--output", default=default_output,
        help="输出目录",
    )
    return parser.parse_args()


if __name__ == "__main__":
    if "-i" in sys.argv or "--input" in sys.argv:
        args = _parse_args()
        extract_mechanical(args.input, args.output)
    else:
        MechApp().mainloop()
