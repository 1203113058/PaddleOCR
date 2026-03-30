"""
机械性能提取 — 共享工具函数
"""

import re

import numpy as np


def clean_cell_text(text: str) -> str:
    """清理单元格文本，多值数字用 / 分隔，OCR 空格断裂自动合并。"""
    text = text.strip()
    if re.match(r'^[\d\s./\-\+%]+$', text):
        space_parts = re.split(r'\s+', text)
        if (len(space_parts) >= 2
                and all(re.match(r'^\d{2,}\.?\d*$', p) for p in space_parts)):
            return "/".join(space_parts)
        squeezed = re.sub(r'\s+', '', text)
        split = _try_split_concat_numbers(squeezed)
        return split if split else squeezed
    return re.sub(r'\s{2,}', ' ', text)


def _try_split_concat_numbers(s: str) -> str | None:
    """
    尝试拆分 OCR 粘连的多组数字，如 '107109126' → '107/109/126'。
    仅当纯数字且能均等拆为 2~4 组、每组 2~4 位时才拆分。
    """
    if not s or not s.isdigit():
        return None
    n = len(s)
    for group_len in (3, 2, 4):
        if n % group_len == 0 and n // group_len >= 2:
            parts = [s[i:i+group_len] for i in range(0, n, group_len)]
            if all(10 <= int(p) for p in parts):
                return "/".join(parts)
    return None


def is_skip_value(text: str) -> bool:
    """判断值是否为 "—" 类占位符或 OCR 误读，不参与判断。"""
    t = text.strip()
    if not t:
        return True
    if t in {"—", "-", "一", "–", "─", "/", "无", "N/A", "n/a"}:
        return True
    # OCR 常将空白/横线误读为单个 l / I / | 等
    if len(t) <= 2 and all(c in "lI|" for c in t):
        return True
    return False


def seq_for_ocr(x) -> list:
    """将 Paddle 返回的 ndarray / list / None 统一转 Python list。"""
    if x is None:
        return []
    if isinstance(x, np.ndarray):
        return x.tolist()
    if isinstance(x, (list, tuple)):
        return list(x)
    return [x]


def rec_box_to_xyxy(box) -> list[int] | None:
    """将 rec_boxes/rec_polys 单项转为轴对齐 [x1,y1,x2,y2]。"""
    if box is None:
        return None
    arr = np.asarray(box, dtype=float).reshape(-1)
    if arr.size < 4:
        return None
    if arr.size == 4:
        x1, y1, x2, y2 = (float(arr[i]) for i in range(4))
    else:
        pts = arr.reshape(-1, 2)
        x1, x2 = float(pts[:, 0].min()), float(pts[:, 0].max())
        y1, y2 = float(pts[:, 1].min()), float(pts[:, 1].max())
    return [int(round(x1)), int(round(y1)), int(round(x2)), int(round(y2))]
