"""
三方对比 — 主流程编排

串联两个 PDF 的 OCR 提取流程，输出统一对比 Excel。

调用方式：
    from mechanical.compare_pipeline import run_comparison
    xlsx = run_comparison(
        supplier_pdf  = "锻件质量证明书-3.pdf",
        internal_pdf  = "机械性能复检报告.pdf",
        output_dir    = "output/",
    )
"""

import shutil
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from .compare import build_comparison, write_comparison_excel
from .excel_writer import html_table_to_sheet
from .field_extractor import extract_named_fields
from .internal_extractor import extract_internal_fields
from .ocr_engine import create_engine, extract_tables_from_page
from .pipeline import extract_mechanical
from .preprocess import image_to_pages, pdf_to_images
from .section_filter import (
    extract_mechanical_section_html,
    html_table_text,
    is_mechanical_table,
)

try:
    from ocr_db import ensure_nonconformity_table, insert_nonconformity_records
except ImportError:
    def ensure_nonconformity_table(): pass  # type: ignore[misc]
    def insert_nonconformity_records(rows): pass  # type: ignore[misc]


def _ocr_internal_pdf(
    pdf_path: str,
    out_dir: Path,
    timestamp: str,
    remove_stamp: bool = True,
) -> list[dict]:
    """
    对内部检测报告 PDF 运行 OCR，提取全量结构化字段。
    返回 extract_internal_fields 的结果列表。
    """
    suffix = Path(pdf_path).suffix.lower()
    tmp_dir = out_dir / "_tmp_internal"

    if suffix == ".pdf":
        print(f"\n  [内部检测] 渲染 PDF 页面…")
        pages = pdf_to_images(pdf_path, tmp_dir=tmp_dir, remove_stamp=remove_stamp)
    elif suffix in (".jpg", ".jpeg", ".png"):
        print(f"\n  [内部检测] 读取图片文件…")
        pages = image_to_pages(pdf_path, tmp_dir=tmp_dir, remove_stamp=remove_stamp)
    else:
        print(f"  [内部检测] 不支持的文件格式：{suffix}")
        return []

    engine = create_engine()
    all_internal: list[dict] = []

    for page_no, img_path in pages:
        print(f"  [内部检测] 第 {page_no} 页  OCR 识别中…")
        tables, _ = extract_tables_from_page(engine, img_path)

        for tbl_info in tables:
            html = tbl_info["html"]
            tmp_wb = Workbook()
            tmp_ws = tmp_wb.active
            html_table_to_sheet(tmp_ws, html, row_offset=1)
            fields = extract_internal_fields(tmp_ws)
            if fields:
                print(f"    ✓ 提取到 {len(fields)} 个检测项（含标准值）")
                all_internal.extend(fields)
            else:
                print(f"    — 未提取到有效检测项（跳过）")

    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)

    return all_internal


def run_comparison(
    supplier_pdf: str,
    internal_pdf: str,
    output_dir: str,
    remove_stamp: bool = True,
) -> str | None:
    """
    三方对比主流程。

    参数：
        supplier_pdf:  供应商质量证明书 PDF / 图片路径
        internal_pdf:  内部机械性能检测报告 PDF / 图片路径
        output_dir:    输出目录
        remove_stamp:  是否自动去除红色公章

    返回：
        生成的对比 Excel 路径（失败时返回 None）
    """
    total_start = time.perf_counter()
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    supplier_name = Path(supplier_pdf).stem

    out_dir = Path(output_dir) / f"{supplier_name}_对比报告_{timestamp}"
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'═'*60}")
    print(f"  [三方对比报告]")
    print(f"  供应商文件：{Path(supplier_pdf).name}")
    print(f"  内部检测文件：{Path(internal_pdf).name}")
    print(f"  输出目录：{out_dir}")
    print(f"  时间戳：{timestamp}")
    print(f"{'═'*60}")

    # ── 步骤 1：提取供应商检测值 ─────────────────────────────────
    print(f"\n[1/3] 提取供应商检测值…")
    supplier_result = extract_mechanical(
        supplier_pdf,
        str(out_dir / "_supplier"),
        remove_stamp=remove_stamp,
    )
    supplier_fields: list[dict] = supplier_result.get("fields", [])
    if not supplier_fields:
        print("  [警告] 供应商文件未提取到任何力学字段。")
    else:
        print(f"  ✓ 供应商字段：{[k for k in supplier_fields[0] if k != '方向']}")

    # ── 步骤 2：提取内部检测值 + 标准值 ──────────────────────────
    print(f"\n[2/3] 提取内部检测值和标准值…")
    internal_fields = _ocr_internal_pdf(
        internal_pdf, out_dir, timestamp, remove_stamp=remove_stamp
    )
    if not internal_fields:
        print("  [警告] 内部检测文件未提取到任何检测项。")
    else:
        print(f"  ✓ 内部检测字段：{[item['检测项'] for item in internal_fields]}")

    # ── 步骤 3：对比 + 输出 Excel ─────────────────────────────────
    print(f"\n[3/3] 生成对比报告…")
    if not supplier_fields and not internal_fields:
        print("  [错误] 两个文件均无有效数据，无法生成报告。")
        return None

    comparison_rows = build_comparison(supplier_fields, internal_fields)

    print(f"\n  ── 对比结果预览 ──")
    for row in comparison_rows:
        status = row["是否合格"]
        mark = "✓" if status == "合格" else ("?" if status == "待确认" else "✗")
        print(f"  {mark} {row['检测项']:8s}  标准:{row['标准值']:10s}  "
              f"供应商:{row['供应商数值']:12s}  内部:{row['内部检测值']:12s}  {status}")

    xlsx_path = out_dir / f"对比报告_{supplier_name}_{timestamp}.xlsx"
    write_comparison_excel(comparison_rows, str(xlsx_path))

    # ── 把不合格记录写入数据库 ────────────────────────────────────
    ensure_nonconformity_table()
    db_rows = []
    for r_idx, row in enumerate(comparison_rows, 1):
        if row["是否合格"] != "不合格":
            continue
        supplier_reason = row["供应商不合格原因"] if row["供应商不合格原因"] != "—" else ""
        internal_reason = row["内部不合格原因"] if row["内部不合格原因"] != "—" else ""
        fail_reason = "; ".join(filter(None, [supplier_reason, internal_reason]))
        actual_val = (
            f"供应商:{row['供应商数值']} 内部:{row['内部检测值']}"
        )
        db_rows.append({
            "report_name":    Path(supplier_pdf).name,
            "report_type":    "三方对比",
            "input_full_path": str(Path(supplier_pdf).resolve()),
            "sheet_name":     "三方对比",
            "page_no":        None,
            "table_index":    None,
            "sample_no":      "",
            "test_item":      row["检测项"],
            "standard_text":  row["标准值"] if row["标准值"] != "—" else "",
            "rule_type":      "compare",
            "actual_value":   actual_val[:512],
            "fail_reason":    fail_reason[:1024],
            "excel_row":      r_idx,
            "excel_col":      0,
            "excel_col_letter": "",
            "xlsx_path":      str(xlsx_path),
            "batch_id":       timestamp,
        })
    insert_nonconformity_records(db_rows)

    total_elapsed = time.perf_counter() - total_start
    fail_count = sum(1 for r in comparison_rows if r["是否合格"] == "不合格")
    pending_count = sum(1 for r in comparison_rows if r["是否合格"] == "待确认")

    print(f"\n{'═'*60}")
    print(f"✓ 完成！总耗时：{total_elapsed:.1f}s")
    print(f"  {xlsx_path.name}")
    print(f"  检测项总计：{len(comparison_rows)}  "
          f"不合格：{fail_count}  待确认：{pending_count}")
    print(f"{'═'*60}")

    return str(xlsx_path)
