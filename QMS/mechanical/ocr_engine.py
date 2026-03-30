"""
机械性能提取 — PPStructure V3 表格识别引擎
"""

import re

import numpy as np
from paddleocr import PPStructureV3

from .utils import seq_for_ocr, rec_box_to_xyxy


def _extract_ocr_lines(r: dict) -> list[dict]:
    ocr = r.get("overall_ocr_res")
    if not isinstance(ocr, dict):
        return []
    texts = seq_for_ocr(ocr.get("rec_texts"))
    boxes = seq_for_ocr(ocr.get("rec_boxes"))
    polys = seq_for_ocr(ocr.get("rec_polys"))
    lines: list[dict] = []
    for i, t in enumerate(texts):
        if t is None:
            continue
        bb = None
        if i < len(boxes):
            bb = rec_box_to_xyxy(boxes[i])
        if bb is None and i < len(polys):
            bb = rec_box_to_xyxy(polys[i])
        if bb is None:
            continue
        lines.append({"text": str(t), "bbox": bb})
    return lines


def _filter_ocr_lines_to_table(
    ocr_lines: list[dict], table_bbox: list | None, margin: int = 8,
) -> list[dict]:
    if not ocr_lines:
        return []
    tb = seq_for_ocr(table_bbox) if table_bbox is not None else []
    if len(tb) < 4:
        return list(ocr_lines)
    tx1, ty1, tx2, ty2 = (float(tb[i]) for i in range(4))
    out: list[dict] = []
    for L in ocr_lines:
        bb = L.get("bbox")
        if not bb or len(bb) < 4:
            continue
        cx = (bb[0] + bb[2]) / 2
        cy = (bb[1] + bb[3]) / 2
        if (tx1 - margin) <= cx <= (tx2 + margin) and (ty1 - margin) <= cy <= (ty2 + margin):
            out.append(L)
    return out if out else list(ocr_lines)


def _parse_cell_bboxes(s: str) -> list[list[int]] | None:
    m = re.search(r"cell_bbox[s]?[:\s\t]+(\[\s*\[.+?\]\s*\])", s, re.DOTALL)
    if not m:
        return None
    try:
        raw = m.group(1)
        rows = re.findall(r"\[([^\[\]]+)\]", raw)
        result = []
        for row in rows:
            vals = re.findall(r"-?\d+\.?\d*", row)
            if len(vals) >= 4:
                result.append([int(float(v)) for v in vals[:4]])
        return result if result else None
    except Exception:
        return None


def create_engine(lang: str = "ch", device: str = "cpu") -> PPStructureV3:
    """创建 PPStructureV3 引擎实例。"""
    return PPStructureV3(lang=lang, device=device)


def extract_tables_from_page(
    engine: PPStructureV3, img_path: str,
) -> tuple[list[dict], list[dict]]:
    """
    对单页图片运行 PPStructureV3。

    返回 (表格列表, OCR 文本行列表)。
    表格项含 html / table_bbox / cell_bboxes / table_ocr_lines。
    """
    results = list(engine.predict(img_path))
    if not results:
        return [], []
    r = results[0]
    ocr_lines = _extract_ocr_lines(r)

    tbl_cell_bboxes: list[list[list[int]] | None] = []
    for tbl_res in r.get("table_res_list", []):
        tbl_str = str(tbl_res)
        cell_bboxes = _parse_cell_bboxes(tbl_str)
        if cell_bboxes is None:
            try:
                raw_boxes = tbl_res["cell_box_list"]
                if raw_boxes:
                    bboxes = []
                    for b in raw_boxes:
                        arr = np.asarray(b, dtype=float).reshape(-1)
                        if arr.size >= 4:
                            bboxes.append([int(round(float(arr[0]))),
                                           int(round(float(arr[1]))),
                                           int(round(float(arr[2]))),
                                           int(round(float(arr[3])))])
                    if bboxes:
                        cell_bboxes = bboxes
            except Exception:
                pass
        tbl_cell_bboxes.append(cell_bboxes)

    tables = []
    tbl_idx = 0
    for item in r.get("parsing_res_list", []):
        item_str = str(item)
        if "label:\ttable" not in item_str and "label: table" not in item_str:
            continue
        html = ""
        m = re.search(r"content:\t(.*)", item_str, re.DOTALL)
        if m:
            html = m.group(1).strip()
        if not html:
            continue
        table_bbox = None
        bm = re.search(r"bbox[:\s\t]+\[([^\]]+)\]", item_str)
        if bm:
            try:
                vals = re.findall(r"-?\d+\.?\d*", bm.group(1))
                if len(vals) >= 4:
                    table_bbox = [int(float(v)) for v in vals[:4]]
            except (ValueError, IndexError):
                pass
        cell_bboxes = tbl_cell_bboxes[tbl_idx] if tbl_idx < len(tbl_cell_bboxes) else None
        if isinstance(cell_bboxes, np.ndarray):
            cell_bboxes = cell_bboxes.tolist()
        scoped = _filter_ocr_lines_to_table(ocr_lines, table_bbox)
        tables.append({
            "html": html,
            "table_bbox": table_bbox,
            "cell_bboxes": cell_bboxes,
            "table_ocr_lines": scoped,
        })
        tbl_idx += 1
    return tables, ocr_lines
