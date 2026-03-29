"""
PDF 表格识别 → Excel 输出（PP-Structure V3 版）

与 pdf2excel.py（逐行文字块方案）的对比版本。
使用 PPStructureV3 直接识别表格结构，输出 HTML 后转 Excel，
单元格内多行文字自动合并，colspan/rowspan 自动还原。

用法：
    conda activate paddleocr310
    python pdf2excel_ppstruct.py -i <PDF路径> [-o <输出目录>]

示例：
    python pdf2excel_ppstruct.py -i "/Users/project/QMS/PaddleOCR/文件/机械性能复检报告.pdf"

输出文件：
    <pdf名>_ppstruct_table.xlsx  — 还原表格结构的 Excel（含单元格合并）
    <pdf名>_ppstruct_page_N.html — 每页表格的原始 HTML（供调试）
"""

import argparse
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

os.environ["PADDLE_PDX_MODEL_SOURCE"] = "modelscope"
os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"

import paddle
paddle.set_device("cpu")

import pypdfium2 as pdfium
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import numpy as np
from paddleocr import PPStructureV3
from PIL import Image, ImageDraw

from ocr_config import load_config

MAX_IMAGE_WIDTH = 1600


# ─── 命令行参数 ───────────────────────────────────────────────────────────────

def parse_args():
    cfg = load_config()
    parser = argparse.ArgumentParser(
        description="PDF / 图片表格识别（PP-Structure V3 版）",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "-i", "--input", required=True,
        help="输入文件路径，支持 PDF / JPG / JPEG / PNG",
    )
    parser.add_argument(
        "-o", "--output",
        default=cfg["output"]["output_dir"],
        help="输出目录",
    )
    return parser.parse_args()


# ─── 公章去除 ─────────────────────────────────────────────────────────────────

def remove_red_stamp(img: Image.Image) -> Image.Image:
    """
    将图片中的红色区域（公章印章）替换为白色，减少 OCR 干扰。

    判断条件（RGB）：
      R > 150  且  G < 90  且  B < 90
    适用于常见的红色圆形公章；蓝色/黑色章不受影响。
    """
    arr = np.array(img.convert("RGB"), dtype=np.uint8)
    r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
    mask = (r > 150) & (g < 90) & (b < 90)
    arr[mask] = [255, 255, 255]
    result = Image.fromarray(arr)
    removed = int(mask.sum())
    if removed > 0:
        print(f"  [去章] 已清除红色区域 {removed} 个像素")
    return result


# ─── PDF 渲染 ─────────────────────────────────────────────────────────────────

def pdf_to_images(
    pdf_path: str,
    tmp_dir: str | Path,
    remove_stamp: bool = True,
    max_width: int = MAX_IMAGE_WIDTH,
) -> list[tuple[int, str]]:
    """
    将 PDF 每页渲染为临时图片，返回 [(页码, 图片路径), ...]。
    图片保存到 tmp_dir（临时目录），识别完成后由调用方统一清理。
    remove_stamp=True 时自动去除红色公章。
    """
    pdf      = pdfium.PdfDocument(pdf_path)
    tmp      = Path(tmp_dir)
    tmp.mkdir(parents=True, exist_ok=True)
    pages    = []
    for i in range(len(pdf)):
        page  = pdf[i]
        w, h  = page.get_size()
        scale = min(max_width / w, 2.0)
        img   = page.render(scale=scale).to_pil()
        if remove_stamp:
            img = remove_red_stamp(img)
        p     = tmp / f"_page_{i+1}.png"
        img.save(str(p))
        print(f"  第 {i+1} 页 → {img.size[0]}×{img.size[1]}  准备完成")
        pages.append((i + 1, str(p)))
    return pages


def image_to_pages(
    img_path: str,
    tmp_dir: str | Path,
    remove_stamp: bool = True,
) -> list[tuple[int, str]]:
    """
    将单张图片（JPG/JPEG/PNG）预处理后保存到 tmp_dir（临时目录），
    返回 [(1, 图片路径)]。识别完成后由调用方统一清理。
    remove_stamp=True 时自动去除红色公章。
    """
    tmp  = Path(tmp_dir)
    tmp.mkdir(parents=True, exist_ok=True)
    dest = tmp / "_page_1.png"
    img  = Image.open(img_path).convert("RGB")
    w, h = img.size
    if w > MAX_IMAGE_WIDTH:
        scale = MAX_IMAGE_WIDTH / w
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    if remove_stamp:
        img = remove_red_stamp(img)
    img.save(str(dest))
    print(f"  图片输入 → {img.size[0]}×{img.size[1]}  准备完成")
    return [(1, str(dest))]


# ─── PP-Structure V3 识别 ─────────────────────────────────────────────────────

def _parse_cell_bboxes(s: str) -> list[list[int]] | None:
    """
    从 PPStructureV3 table_res_list 对象字符串中提取单元格精确坐标列表。
    支持两种常见格式：
      cell_bbox: [[x1,y1,x2,y2], ...]
      cell_bboxes: [[x1,y1,x2,y2], ...]
    返回 [[x1,y1,x2,y2], ...] 或 None。
    """
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


def extract_tables_from_page(engine: PPStructureV3, img_path: str) -> list[dict]:
    """
    对单页图片运行 PPStructureV3，返回识别到的所有表格信息列表。
    每项：{
      "html": str,
      "table_bbox": [x1, y1, x2, y2] | None,   整个表格坐标
      "cell_bboxes": [[x1,y1,x2,y2], ...] | None  每个 td 的精确坐标（按 HTML td 顺序）
    }
    优先使用 cell_bboxes 做精确图片标注，不可用时退回均匀网格估算。
    """
    results = list(engine.predict(img_path))
    if not results:
        return []

    r = results[0]

    # 尝试从 table_res_list 提取每个表格的精确单元格坐标
    tbl_cell_bboxes: list[list[list[int]] | None] = []
    for tbl_res in r.get("table_res_list", []):
        tbl_str = str(tbl_res)
        cell_bboxes = _parse_cell_bboxes(tbl_str)
        # 也尝试直接访问属性（不同版本可能支持）
        if cell_bboxes is None:
            for attr in ("cell_bbox", "cell_bboxes"):
                val = getattr(tbl_res, attr, None)
                if val is not None:
                    try:
                        cell_bboxes = [[int(v) for v in b[:4]] for b in val]
                    except Exception:
                        pass
                    break
        tbl_cell_bboxes.append(cell_bboxes)

    tables = []
    tbl_idx = 0
    for item in r.get("parsing_res_list", []):
        item_str = str(item)
        if "label:\ttable" in item_str or "label: table" in item_str:
            html = ""
            m = re.search(r"content:\t(.*)", item_str, re.DOTALL)
            if m:
                html = m.group(1).strip()
            if not html:
                continue

            # 整个表格的外框 bbox
            table_bbox = None
            bm = re.search(r"bbox[:\s\t]+\[([^\]]+)\]", item_str)
            if bm:
                try:
                    vals = re.findall(r"-?\d+\.?\d*", bm.group(1))
                    if len(vals) >= 4:
                        table_bbox = [int(float(v)) for v in vals[:4]]
                except (ValueError, IndexError):
                    pass

            # 单元格精确坐标
            cell_bboxes = tbl_cell_bboxes[tbl_idx] if tbl_idx < len(tbl_cell_bboxes) else None

            if cell_bboxes:
                print(f"    → 表格 cell_bboxes：{len(cell_bboxes)} 个单元格坐标（精确模式）")
            elif table_bbox:
                print(f"    → 表格 bbox：{table_bbox}（均匀估算模式）")
            else:
                print(f"    → 表格坐标：未提取到（图片标注将跳过）")

            tables.append({"html": html, "table_bbox": table_bbox, "cell_bboxes": cell_bboxes})
            tbl_idx += 1
    return tables


# ─── HTML 表格 → openpyxl Sheet ───────────────────────────────────────────────

# ─── 判定规则引擎 ─────────────────────────────────────────────────────────────

def parse_requirement(req: str) -> dict | None:
    """
    解析要求值字符串，返回判定规则字典。

    支持四种形式：
      区间      686-800   → {"type": "range", "min": 686, "max": 800}
      大于等于  ≥15       → {"type": "gte",   "value": 15}
      小于等于  ≤X        → {"type": "lte",   "value": X}
      精确等于  -50       → {"type": "eq",    "value": -50}
    """
    req = req.strip()
    if not req:
        return None
    # 区间：两个正数用 - 分隔（负数不会匹配，因为负数开头没有纯数字）
    m = re.match(r'^(\d+\.?\d*)\s*[-–]\s*(\d+\.?\d*)$', req)
    if m:
        return {"type": "range", "min": float(m.group(1)), "max": float(m.group(2))}
    # 大于等于：≥ 或 >=
    m = re.match(r'^[≥>]=?\s*(-?\d+\.?\d*)$', req)
    if m:
        return {"type": "gte", "value": float(m.group(1))}
    # 小于等于：≤ 或 <=
    m = re.match(r'^[≤<]=?\s*(-?\d+\.?\d*)$', req)
    if m:
        return {"type": "lte", "value": float(m.group(1))}
    # 精确等于：纯数字或负数
    m = re.match(r'^-?\d+\.?\d*$', req)
    if m:
        return {"type": "eq", "value": float(req)}
    return None


def check_value(actual: str, rule: dict) -> bool | None:
    """
    判断实测值是否满足要求规则。
    返回 True=合格 / False=不合格 / None=无法解析（需人工复核）

    支持多值：冲击吸收能量如 "120/124 /124"，以 / 或空格分隔，
    所有子值均需满足要求才判定为合格。
    """
    actual = actual.strip()
    if not actual:
        return None

    parts = re.split(r'[/\s]+', actual)
    values: list[float] = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        try:
            values.append(float(p))
        except ValueError:
            return None

    if not values:
        return None

    def _ok(v: float) -> bool:
        t = rule["type"]
        if t == "range":
            return rule["min"] <= v <= rule["max"]
        if t == "gte":
            return v >= rule["value"]
        if t == "lte":
            return v <= rule["value"]
        if t == "eq":
            return abs(v - rule["value"]) < 0.01
        return True

    return all(_ok(v) for v in values)


def apply_value_judgment(ws) -> list[tuple[int, int, str, str]]:
    """
    遍历已写入的 worksheet，自动找到「要求值」行和「试样」行，
    按列对比实测值与要求值，并对单元格着色：
      - 不合格：红色背景 #FFCCCC + 红色加粗字体
      - 无法解析：黄色背景 #FFF2CC（需人工复核）

    返回：不合格单元格列表 [(excel_row, excel_col, actual_text, reason), ...]
    """
    FILL_FAIL    = PatternFill(fill_type="solid", fgColor="FFCCCC")
    FILL_UNKNOWN = PatternFill(fill_type="solid", fgColor="FFF2CC")
    FONT_FAIL    = Font(bold=True, color="CC0000")

    max_row = ws.max_row
    max_col = ws.max_column

    # 构建行列文本矩阵（合并单元格非主格值为空串）
    rows_data: list[tuple[int, list[str]]] = []
    for r in range(1, max_row + 1):
        vals = []
        for c in range(1, max_col + 1):
            raw = ws.cell(row=r, column=c).value
            vals.append(str(raw).strip() if raw is not None else "")
        rows_data.append((r, vals))

    # 定位「要求值」行
    req_excel_row: int | None = None
    req_vals: list[str] = []
    for excel_row, row_vals in rows_data:
        if any("要求值" in v or "Standard" in v for v in row_vals):
            req_excel_row = excel_row
            req_vals = row_vals
            break

    if req_excel_row is None:
        print("  [判定] 未找到「要求值」行，跳过合规判定。")
        return []

    # 预解析每列的要求规则（列索引 0-based）
    col_rules: dict[int, dict] = {}
    for col_idx, req in enumerate(req_vals):
        rule = parse_requirement(req)
        if rule:
            col_rules[col_idx] = rule

    if not col_rules:
        print("  [判定] 要求值行中未解析到有效规则，跳过。")
        return []

    # 定位试样行（第一个非空值符合 "数字-数字" 格式，如 22-9637）
    judged = 0
    failures: list[tuple[int, int, str, str]] = []
    for excel_row, row_vals in rows_data:
        if excel_row <= req_excel_row:
            continue
        first_val = next((v for v in row_vals if v), "")
        if not re.match(r'^\d{2,}[-–]\d{4,}', first_val):
            continue

        # 对有规则的列逐一判定
        for col_idx, rule in col_rules.items():
            actual = row_vals[col_idx] if col_idx < len(row_vals) else ""
            if not actual:
                continue

            result = check_value(actual, rule)
            cell = ws.cell(row=excel_row, column=col_idx + 1)

            if result is False:
                cell.fill = FILL_FAIL
                cell.font = FONT_FAIL
                t = rule["type"]
                if t == "range":
                    reason = f"{actual}<{rule['min']}或>{rule['max']}"
                elif t == "gte":
                    reason = f"{actual}<{rule['value']}"
                elif t == "lte":
                    reason = f"{actual}>{rule['value']}"
                else:
                    reason = f"{actual}≠{rule['value']}"
                failures.append((excel_row, col_idx + 1, actual, reason))
                judged += 1
            elif result is None:
                cell.fill = FILL_UNKNOWN

    print(f"  [判定] 完成，不合格单元格：{judged} 个（红色标注）。")
    return failures


def _col_idx(col: int) -> str:
    return get_column_letter(col)


def clean_cell_text(text: str) -> str:
    """
    清理单元格文字中的多余空格。

    - 数字型内容（仅含数字、小数点、负号、斜杠、百分号）：去除所有内部空格
      例：'7 26' → '726'，'1 20/124 /124' → '120/124/124'
    - 文字型内容：将连续多个空格合并为一个
    """
    text = text.strip()
    if re.match(r'^[\d\s./\-\+%]+$', text):
        return re.sub(r'\s+', '', text)
    return re.sub(r'\s{2,}', ' ', text)


def html_table_to_sheet(
    ws, html: str, row_offset: int = 1
) -> tuple[int, dict[tuple[int, int], int]]:
    """
    将 HTML <table> 写入 openpyxl worksheet。
    - 处理 colspan / rowspan 单元格合并
    - 返回 (最后行号, td_coord_map)
      td_coord_map: {(excel_row, excel_col): td_index}
      td_index 与 cell_bboxes 列表顺序一致，用于精确图片标注。

    row_offset: 从第几行开始写（用于多页拼接时追加）
    """
    soup  = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return row_offset, {}

    occupied: dict[tuple[int, int], bool] = {}
    td_coord_map: dict[tuple[int, int], int] = {}
    td_index = 0

    excel_row = row_offset
    for tr in table.find_all("tr"):
        excel_col = 1
        for td in tr.find_all(["td", "th"]):
            while occupied.get((excel_row, excel_col)):
                excel_col += 1

            colspan = int(td.get("colspan", 1))
            rowspan = int(td.get("rowspan", 1))
            text    = clean_cell_text(td.get_text(separator=" ", strip=True))

            ws.cell(row=excel_row, column=excel_col, value=text)

            # 记录该 td 对应的 Excel 坐标和 td 顺序索引
            td_coord_map[(excel_row, excel_col)] = td_index
            td_index += 1

            if colspan > 1 or rowspan > 1:
                end_row = excel_row + rowspan - 1
                end_col = excel_col + colspan - 1
                ws.merge_cells(
                    start_row=excel_row, start_column=excel_col,
                    end_row=end_row,     end_column=end_col,
                )
                for r in range(excel_row, end_row + 1):
                    for c in range(excel_col, end_col + 1):
                        if r != excel_row or c != excel_col:
                            occupied[(r, c)] = True

            cell = ws.cell(row=excel_row, column=excel_col)
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

            excel_col += colspan

        excel_row += 1

    return excel_row, td_coord_map


def apply_header_style(ws, max_row: int, max_col: int):
    """对第一行加粗、浅灰背景。"""
    fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
    for col in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True)
        cell.fill = fill


def auto_column_width(ws):
    """自动调整列宽（按最长内容估算）。"""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len * 1.2, 8), 50)


# ─── 图片标注 ──────────────────────────────────────────────────────────────────

def _get_cell_img_bbox(
    tbl_info: dict,
    td_coord_map: dict,
    excel_row: int,
    excel_col: int,
) -> list[int] | None:
    """
    获取单元格在图片上的精确坐标。
    优先使用 cell_bboxes（精确），不可用时退回均匀网格估算。
    返回 [ix1, iy1, ix2, iy2] 或 None。
    """
    cell_bboxes = tbl_info.get("cell_bboxes")
    td_idx      = td_coord_map.get((excel_row, excel_col))

    # 精确模式：用 PPStructureV3 提供的单元格坐标
    if cell_bboxes and td_idx is not None and td_idx < len(cell_bboxes):
        return cell_bboxes[td_idx]

    # 退回模式：均匀网格估算
    table_bbox = tbl_info.get("table_bbox")
    html       = tbl_info.get("html", "")
    if not table_bbox or not html:
        return None

    soup  = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return None

    rows   = table.find_all("tr")
    n_rows = len(rows)
    n_cols = max(
        (sum(int(td.get("colspan", 1)) for td in tr.find_all(["td", "th"])) for tr in rows),
        default=0,
    )
    if n_rows == 0 or n_cols == 0:
        return None

    tx1, ty1, tx2, ty2 = table_bbox
    cell_w = (tx2 - tx1) / n_cols
    cell_h = (ty2 - ty1) / n_rows
    r_idx  = excel_row - 1
    c_idx  = excel_col - 1
    return [
        int(tx1 + c_idx * cell_w),
        int(ty1 + r_idx * cell_h),
        int(tx1 + (c_idx + 1) * cell_w),
        int(ty1 + (r_idx + 1) * cell_h),
    ]


def mark_failures_on_image(
    img_path: str,
    failures_with_meta: list,
    out_path: str,
    td_coord_maps: dict | None = None,
) -> None:
    """
    在 PDF 渲染图片上以红色矩形框标注所有不合格单元格，并在框旁标注原因。

    failures_with_meta: [(tbl_info_dict, excel_row, excel_col, actual, reason), ...]
    td_coord_maps: {tbl_info_id: td_coord_map}，由 html_table_to_sheet 返回，用于精确定位
    out_path: 保存标注后图片的路径（jpg）
    """
    img   = Image.open(img_path).convert("RGB")
    draw  = ImageDraw.Draw(img)
    drawn = 0

    for tbl_info, excel_row, excel_col, actual, reason in failures_with_meta:
        tbl_id       = id(tbl_info)
        td_coord_map = (td_coord_maps or {}).get(tbl_id, {})
        cell_bbox    = _get_cell_img_bbox(tbl_info, td_coord_map, excel_row, excel_col)

        if cell_bbox is None:
            print(f"  [跳过] 行{excel_row} 列{excel_col} 无法获取图片坐标，跳过标注。")
            continue

        ix1, iy1, ix2, iy2 = cell_bbox
        draw.rectangle([ix1, iy1, ix2, iy2], outline=(220, 50, 50), width=3)

        label  = reason or actual
        text_x = min(ix2 + 4, img.width - 1)
        text_y = max(iy1, 0)
        draw.text((text_x, text_y), label, fill=(220, 50, 50))
        drawn += 1

    if drawn > 0:
        img.save(out_path, quality=95)
        print(f"  标注图片已保存：{Path(out_path).name}（共标注 {drawn} 处）")
    else:
        print(f"  [跳过] 未画出任何标注框，图片未保存。")


# ─── 主流程 ───────────────────────────────────────────────────────────────────

def ocr_pdf_ppstruct(
    pdf_path: str,
    output_dir: str,
    report_type: str = "力学检测",
    remove_stamp: bool = True,
):
    """
    report_type:   报告类型，用于输出文件命名前缀（"力学检测"/"化学检测" 等）
    remove_stamp:  True 时自动去除图片/PDF 中的红色公章，减少 OCR 干扰
    输出文件命名规则：
      图片_{report_type}_{编号}_{时间戳}.png       — PDF 渲染原图（已去章）
      错误标注_{编号}_{时间戳}.jpg                 — 不合格标注图
      表格_{report_type}_{时间戳}.xlsx             — 识别结果表格
    """
    from collections import defaultdict

    pdf_name    = Path(pdf_path).stem
    total_start = time.perf_counter()
    timestamp   = datetime.now().strftime("%Y%m%d%H%M%S")

    # 在指定输出目录下，以"文件名_识别结果_时间戳"新建子文件夹
    out_dir = Path(output_dir) / f"{pdf_name}_识别结果_{timestamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'═'*55}")
    print(f"  [PP-Structure V3]  {Path(pdf_path).name}")
    print(f"  报告类型：{report_type}")
    print(f"  输出目录：{out_dir}")
    print(f"  时间戳：{timestamp}")
    print(f"{'═'*55}")

    # ── 1. 准备图片（PDF 渲染 或 直接使用图片）────────────────────────────────
    # 临时目录：存放渲染中间图片，识别完成后自动删除
    tmp_dir = out_dir / "_tmp"
    suffix  = Path(pdf_path).suffix.lower()
    if suffix == ".pdf":
        print(f"\n[1/4] 渲染 PDF 页面{'（自动去章）' if remove_stamp else ''}…")
        pages = pdf_to_images(pdf_path, tmp_dir=tmp_dir, remove_stamp=remove_stamp)
    elif suffix in (".jpg", ".jpeg", ".png"):
        print(f"\n[1/4] 读取图片文件{'（自动去章）' if remove_stamp else ''}…")
        pages = image_to_pages(pdf_path, tmp_dir=tmp_dir, remove_stamp=remove_stamp)
    else:
        print(f"  [错误] 不支持的文件格式：{suffix}，仅支持 PDF / JPG / JPEG / PNG")
        return

    # ── 2. PPStructureV3 识别 ─────────────────────────────────────────────────
    print("\n[2/4] 启动 PPStructureV3…")
    engine = PPStructureV3(lang="ch", device="cpu")

    wb  = Workbook()
    # all_tables: [(页码, 表格序号, tbl_info_dict)]
    all_tables: list[tuple[int, int, dict]] = []

    for page_no, img_path in pages:
        t0 = time.perf_counter()
        print(f"\n  第 {page_no} 页  识别中…")
        tables = extract_tables_from_page(engine, img_path)
        elapsed = time.perf_counter() - t0
        print(f"  第 {page_no} 页  检测到 {len(tables)} 个表格  耗时 {elapsed:.1f}s")

        for j, tbl_info in enumerate(tables, 1):
            all_tables.append((page_no, j, tbl_info))
            # 保存原始 HTML 供调试
            html_path = out_dir / f"{pdf_name}_ppstruct_page_{page_no}_table_{j}.html"
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(f"<html><meta charset='utf-8'><body>{tbl_info['html']}</body></html>")
            print(f"    → HTML 已保存：{html_path.name}")

    # ── 3. 写入 Excel ─────────────────────────────────────────────────────────
    print(f"\n[3/4] 写入 Excel…  共 {len(all_tables)} 个表格")

    if not all_tables:
        print("  未检测到任何表格，退出。")
        return

    # page_failures_map: page_no → [(tbl_info, excel_row, excel_col, actual, reason)]
    page_failures_map: dict[int, list] = defaultdict(list)
    # td_coord_maps: tbl_info id → td_coord_map（用于精确图片标注）
    td_coord_maps: dict[int, dict] = {}

    for idx, (page_no, tbl_no, tbl_info) in enumerate(all_tables):
        sheet_name = f"P{page_no}_T{tbl_no}" if len(all_tables) > 1 else "Sheet1"
        ws = wb.active if idx == 0 else wb.create_sheet(sheet_name)
        if idx == 0:
            ws.title = sheet_name
        last_row, td_coord_map = html_table_to_sheet(ws, tbl_info["html"], row_offset=1)
        td_coord_maps[id(tbl_info)] = td_coord_map
        max_col  = ws.max_column
        apply_header_style(ws, last_row, max_col)
        failures = apply_value_judgment(ws)
        auto_column_width(ws)
        print(f"  Sheet [{sheet_name}]  {last_row-1} 行 × {max_col} 列")

        for f_row, f_col, actual, reason in failures:
            page_failures_map[page_no].append((tbl_info, f_row, f_col, actual, reason))

    xlsx_path = out_dir / f"表格_{report_type}_{timestamp}.xlsx"
    wb.save(str(xlsx_path))

    # ── 4. 生成不合格标注图片 ─────────────────────────────────────────────────
    print(f"\n[4/4] 生成不合格标注图片…")
    marked_count = 0
    for page_no, img_path in pages:
        failures_this_page = page_failures_map.get(page_no, [])
        if not failures_this_page:
            print(f"  第 {page_no} 页  无不合格项，跳过图片标注。")
            continue
        marked_path = out_dir / f"错误标注_{page_no}_{timestamp}.jpg"
        mark_failures_on_image(img_path, failures_this_page, str(marked_path), td_coord_maps)
        marked_count += 1

    # 清理临时图片目录
    import shutil
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
        print("  [清理] 临时图片已删除。")

    total_elapsed = time.perf_counter() - total_start
    print(f"\n{'═'*55}")
    print(f"✓ 完成！总耗时：{total_elapsed:.1f}s")
    print(f"  {xlsx_path.name}  — 识别结果表格（含合规标注）")
    if marked_count:
        print(f"  错误标注_N_{timestamp}.jpg  — 不合格标注图片（共 {marked_count} 页）")
    print(f"{'═'*55}")


# ─── stdout 重定向 ────────────────────────────────────────────────────────────

import queue
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


class _QueueWriter:
    """将 stdout/stderr 写入线程安全队列，供 GUI 主线程消费。"""

    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, text: str):
        if text:
            self._q.put(text)

    def flush(self):
        pass


# ─── GUI 主界面 ───────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 表格识别工具（PP-Structure V3）")
        self.minsize(720, 560)
        self.resizable(True, True)

        self._output_dir    = tk.StringVar(value="/Users/project/QMS/PaddleOCR/QMS/output2")
        self._remove_stamp  = tk.BooleanVar(value=True)
        self._pdf_files: list[str] = [
            "/Users/project/QMS/PaddleOCR/QMS/test_files/机械性能复检报告.pdf"
        ]
        self._log_queue: queue.Queue = queue.Queue()
        self._running = False

        self._build_ui()
        self._poll_log()
        self._refresh_files_text()

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

        # ── 操作按钮与状态 ──
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=12, pady=(8, 2))

        self._btn_start = ttk.Button(frm_btn, text="▶  开始识别", command=self._start, width=14)
        self._btn_start.pack(side="left", padx=(0, 8))

        self._btn_stop = ttk.Button(frm_btn, text="■  停止", command=self._stop,
                                    state="disabled", width=10)
        self._btn_stop.pack(side="left", padx=(0, 12))

        ttk.Checkbutton(
            frm_btn, text="自动去除公章", variable=self._remove_stamp,
        ).pack(side="left", padx=(0, 12))

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

    # ─── 文件 / 目录选择 ──────────────────────────────────────────────────────

    def _pick_files(self):
        files = filedialog.askopenfilenames(
            title="选择 PDF 或图片文件",
            filetypes=[
                ("支持的格式", "*.pdf *.jpg *.jpeg *.png"),
                ("PDF 文件", "*.pdf"),
                ("图片文件", "*.jpg *.jpeg *.png"),
                ("所有文件", "*.*"),
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
        raw  = self._output_dir.get().strip()
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
        self._status_var.set("正在等待当前文件完成后停止…")

    def _run_ocr(self):
        old_stdout, old_stderr = sys.stdout, sys.stderr
        writer = _QueueWriter(self._log_queue)
        sys.stdout = writer
        sys.stderr = writer

        try:
            total     = len(self._pdf_files)
            completed = 0
            for idx, pdf_path in enumerate(self._pdf_files, 1):
                if not self._running:
                    print("\n[停止] 已中止识别。")
                    break
                print(f"\n{'═'*50}")
                print(f"[{idx}/{total}]  {Path(pdf_path).name}")
                print(f"{'═'*50}")
                try:
                    ocr_pdf_ppstruct(
                        pdf_path,
                        self._output_dir.get().strip(),
                        remove_stamp=self._remove_stamp.get(),
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
    if "-i" in sys.argv or "--input" in sys.argv:
        args = parse_args()
        ocr_pdf_ppstruct(args.input, args.output)
    else:
        App().mainloop()
