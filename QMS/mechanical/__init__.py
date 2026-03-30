"""
mechanical — 机械性能（力学试验）表格专项提取包

核心接口:
    extract_mechanical(pdf_path, output_dir, remove_stamp=True) -> dict
    MechApp  — Tkinter GUI 应用类

示例:
    from mechanical import extract_mechanical
    result = extract_mechanical("报告.pdf", "output/")
    print(result["fields"])
"""

import os
os.environ.setdefault("PADDLE_PDX_MODEL_SOURCE", "modelscope")
os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")

import paddle
paddle.set_device("cpu")

from .pipeline import extract_mechanical
from .compare_pipeline import run_comparison
from .gui import MechApp

__all__ = ["extract_mechanical", "run_comparison", "MechApp"]
