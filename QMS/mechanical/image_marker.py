"""
机械性能提取 — 不合格项图片标注
"""

import re
from pathlib import Path

import numpy as np
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw

from .utils import clean_cell_text, seq_for_ocr


def _table_grid_metrics(tbl_info):
    table_bbox = tbl_info.get("table_bbox")
    html = tbl_info.get("html", "")
    tb_list = seq_for_ocr(table_bbox) if table_bbox is not None else []
    if len(tb_list) < 4 or not html:
        return None
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return None
    rows = table.find_all("tr")
    n_rows = len(rows)
    n_cols = max(
        (sum(int(td.get("colspan", 1)) for td in tr.find_all(["td", "th"])) for tr in rows),
        default=0,
    )
    if n_rows == 0 or n_cols == 0:
        return None
    tx1, ty1, tx2, ty2 = (float(tb_list[i]) for i in range(4))
    return tx1, ty1, tx2, ty2, (tx2 - tx1) / n_cols, (ty2 - ty1) / n_rows, n_rows, n_cols


def _grid_rect_pixels(tbl_info, min_row, max_row, min_col, max_col):
    m = _table_grid_metrics(tbl_info)
    if not m:
        return None
    tx1, ty1, _, _, cell_w, cell_h, n_rows, n_cols = m
    r1 = max(0, min(min_row - 1, n_rows - 1))
    r2 = max(0, min(max_row - 1, n_rows - 1))
    c1 = max(0, min(min_col - 1, n_cols - 1))
    c2 = max(0, min(max_col - 1, n_cols - 1))
    if r2 < r1: r1, r2 = r2, r1
    if c2 < c1: c1, c2 = c2, c1
    return [int(tx1 + c1 * cell_w), int(ty1 + r1 * cell_h),
            int(tx1 + (c2 + 1) * cell_w), int(ty1 + (r2 + 1) * cell_h)]


def _grid_cell_bbox(tbl_info, excel_row, excel_col):
    return _grid_rect_pixels(tbl_info, excel_row, excel_row, excel_col, excel_col)


def _count_html_td_cells(html: str) -> int:
    if not html:
        return 0
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return 0
    return sum(len(tr.find_all(["td", "th"])) for tr in table.find_all("tr"))


def _align_cell_bbox_index(td_idx, cell_bboxes, n_td):
    n_b = len(cell_bboxes)
    if n_td <= 0 or td_idx < 0:
        return None
    delta = n_b - n_td
    adj = td_idx + (1 if delta == 1 else (-1 if delta == -1 and td_idx > 0 else 0))
    return adj if 0 <= adj < n_b else None


def _get_cell_img_bbox(tbl_info, td_coord_map, excel_row, excel_col):
    cell_bboxes = tbl_info.get("cell_bboxes")
    td_idx = td_coord_map.get((excel_row, excel_col))
    html = tbl_info.get("html", "")
    if cell_bboxes is not None and len(cell_bboxes) > 0 and td_idx is not None:
        n_td = _count_html_td_cells(html)
        adj = _align_cell_bbox_index(td_idx, cell_bboxes, n_td) if n_td > 0 else (
            td_idx if 0 <= td_idx < len(cell_bboxes) else None)
        if adj is not None:
            bb = cell_bboxes[adj]
            arr = np.asarray(bb, dtype=float).reshape(-1)
            if arr.size >= 4:
                return [int(round(arr[i])) for i in range(4)]
    return _grid_cell_bbox(tbl_info, excel_row, excel_col)


def _bbox_intersection_area(a, b):
    return max(0, min(a[2], b[2]) - max(a[0], b[0])) * max(0, min(a[3], b[3]) - max(a[1], b[1]))


def _expand_bbox(bb, px, img_w, img_h):
    return [max(0, bb[0] - px), max(0, bb[1] - px),
            min(img_w, bb[2] + px), min(img_h, bb[3] + px)]


def _ocr_text_matches(actual, ocr_text):
    a, o = clean_cell_text(actual), clean_cell_text(ocr_text)
    if not a:
        return False
    if a == o:
        return True
    def sq(s):
        return re.sub(r"\s+", "", s.replace("．", ".").replace("，", ","))
    return sq(a) == sq(o)


def _ocr_lines_for_actual(actual, ocr_lines):
    if not actual or not ocr_lines:
        return []
    m = [L for L in ocr_lines if _ocr_text_matches(actual, L["text"])]
    if m:
        return m
    sa = clean_cell_text(str(actual).strip())
    if len(sa) < 2:
        return []
    def sq(s):
        return re.sub(r"\s+", "", s.replace("．", ".").replace("，", ","))
    sqa = sq(sa)
    return [L for L in ocr_lines if sqa in sq(str(L.get("text", ""))) or sq(str(L.get("text", ""))) in sqa]


def _union_bboxes(bboxes):
    return [min(b[0] for b in bboxes), min(b[1] for b in bboxes),
            max(b[2] for b in bboxes), max(b[3] for b in bboxes)]


def _try_merge_split(actual, matches):
    def sq(s):
        return re.sub(r"\s+", "", s.replace("．", ".").replace("，", ","))
    sqa = sq(clean_cell_text(str(actual).strip()))
    if not sqa or len(matches) < 2:
        return None
    sorted_m = sorted(matches, key=lambda L: (L["bbox"][1], L["bbox"][0]))
    if sq("".join(str(L.get("text", "")) for L in sorted_m)) == sqa:
        return _union_bboxes([L["bbox"] for L in sorted_m])
    return None


def _pick_ocr_bbox(actual, cell_bbox, ocr_lines, img_w=None, img_h=None, expand_px=36):
    if not ocr_lines or not actual:
        return None
    matches = _ocr_lines_for_actual(actual, ocr_lines)
    if not matches:
        return None
    if len(matches) == 1:
        return matches[0]["bbox"]
    merged = _try_merge_split(actual, matches)
    if merged is not None:
        return merged
    ex = None
    coarse_h = 0.0
    if cell_bbox is not None and img_w is not None and img_h is not None:
        ex = _expand_bbox(cell_bbox, expand_px, img_w, img_h)
        coarse_h = float(cell_bbox[3] - cell_bbox[1])
    ccy = ccx = 0.0
    if cell_bbox is not None:
        ccy, ccx = (cell_bbox[1] + cell_bbox[3]) / 2, (cell_bbox[0] + cell_bbox[2]) / 2

    def key(L):
        bb = L["bbox"]
        ia = _bbox_intersection_area(bb, ex) if ex else (_bbox_intersection_area(bb, cell_bbox) if cell_bbox else 0)
        bh = float(bb[3] - bb[1])
        area = float((bb[2] - bb[0]) * bh)
        tall_pen = (1e6 + bh) if (coarse_h > 6 and bh > coarse_h * 2.15) else 0.0
        return (ia, -tall_pen, -area, -abs((bb[1]+bb[3])/2 - ccy), -abs((bb[0]+bb[2])/2 - ccx))

    return max(matches, key=key)["bbox"]


def mark_failures_on_image(
    img_path, failures_with_meta, out_path,
    td_coord_maps=None, ocr_lines=None,
):
    """在底图上用红框标注不合格单元格。"""
    img = Image.open(img_path).convert("RGB")
    draw = ImageDraw.Draw(img)
    drawn = 0
    W, H = img.size

    for tbl_info, excel_row, excel_col, actual, reason in failures_with_meta:
        td_coord_map = (td_coord_maps or {}).get(id(tbl_info), {})
        coarse = _grid_cell_bbox(tbl_info, excel_row, excel_col)
        scoped = tbl_info.get("table_ocr_lines") or []
        prefer = scoped if scoped else (ocr_lines or [])
        ocr_bbox = _pick_ocr_bbox(actual, coarse, prefer, W, H)
        if ocr_bbox is None:
            ocr_bbox = _pick_ocr_bbox(actual, coarse, ocr_lines, W, H)

        if ocr_bbox is not None and coarse is not None:
            cx = (ocr_bbox[0] + ocr_bbox[2]) / 2.0
            cw = coarse[2] - coarse[0]
            draw_bbox = [max(0, int(cx - cw/2)), coarse[1], min(W, int(cx + cw/2)), coarse[3]]
        elif ocr_bbox is not None:
            draw_bbox = ocr_bbox
        else:
            draw_bbox = _get_cell_img_bbox(tbl_info, td_coord_map, excel_row, excel_col) or coarse

        if draw_bbox is None:
            continue
        draw.rectangle(draw_bbox, outline=(220, 50, 50), width=3)
        draw.text((min(draw_bbox[2]+4, W-1), max(draw_bbox[1], 0)),
                  reason or actual, fill=(220, 50, 50))
        drawn += 1

    if drawn > 0:
        img.save(out_path, quality=95)
        print(f"  标注图片已保存：{Path(out_path).name}（共标注 {drawn} 处）")
    else:
        print(f"  [跳过] 未画出任何标注框，图片未保存。")
