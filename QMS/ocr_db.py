"""
将 PPStructure 判定产生的不合格记录写入 MySQL。

依赖：pip install pymysql
配置：ocr_config.load_config() 中 [mysql] 节，或环境变量覆盖。
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from ocr_config import load_config

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS `ocr_nonconformity_record` (
  `id`              BIGINT UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '主键',
  `report_name`     VARCHAR(512)    NOT NULL DEFAULT '' COMMENT '报告文件名（不含路径）',
  `report_type`     VARCHAR(64)     NOT NULL DEFAULT '' COMMENT '报告类型：力学检测等',
  `input_full_path` VARCHAR(1024)   NOT NULL DEFAULT '' COMMENT '原始 PDF/图片完整路径',
  `sheet_name`      VARCHAR(128)    NOT NULL DEFAULT '' COMMENT 'Excel 工作表名',
  `page_no`         INT             NULL COMMENT '来源 PDF 页码',
  `table_index`     INT             NULL COMMENT '该页内表格序号',
  `sample_no`       VARCHAR(128)    NOT NULL DEFAULT '' COMMENT '试样编号',
  `test_item`       VARCHAR(768)    NOT NULL DEFAULT '' COMMENT '检测项目（表头推断）',
  `standard_text`   VARCHAR(512)    NOT NULL DEFAULT '' COMMENT '标准/要求值（单元格原文）',
  `rule_type`       VARCHAR(16)     NOT NULL DEFAULT '' COMMENT '判定规则：range/gte/lte/eq',
  `actual_value`    VARCHAR(512)    NOT NULL DEFAULT '' COMMENT '实测值',
  `fail_reason`     VARCHAR(1024)   NOT NULL DEFAULT '' COMMENT '不合格原因',
  `excel_row`       INT             NULL COMMENT 'Excel 行号',
  `excel_col`       INT             NULL COMMENT 'Excel 列号',
  `excel_col_letter` VARCHAR(8)     NOT NULL DEFAULT '' COMMENT 'Excel 列字母',
  `xlsx_path`       VARCHAR(1024)   NOT NULL DEFAULT '' COMMENT '输出 Excel 路径',
  `batch_id`        VARCHAR(64)     NOT NULL DEFAULT '' COMMENT '同次识别批次（时间戳）',
  `created_at`      DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
  `updated_at`      DATETIME        NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  PRIMARY KEY (`id`),
  KEY `idx_report_batch` (`batch_id`),
  KEY `idx_report_name` (`report_name`(191)),
  KEY `idx_created` (`created_at`),
  KEY `idx_sample` (`sample_no`(64))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
  COMMENT='OCR 表格识别-不合格项记录';
"""


def _mysql_params() -> dict[str, Any]:
    cfg = load_config().get("mysql", {})
    return {
        "enabled": bool(cfg.get("enabled", False)),
        "host": os.environ.get("OCR_DB_HOST", cfg.get("host", "127.0.0.1")),
        "port": int(os.environ.get("OCR_DB_PORT", cfg.get("port", 3306))),
        "user": os.environ.get("OCR_DB_USER", cfg.get("user", "root")),
        "password": os.environ.get("OCR_DB_PASSWORD", cfg.get("password", "")),
        "database": os.environ.get("OCR_DB_NAME", cfg.get("database", "sg_local")),
        "table": cfg.get("table", "ocr_nonconformity_record"),
    }


def _connect():
    import pymysql

    p = _mysql_params()
    return pymysql.connect(
        host=p["host"],
        port=p["port"],
        user=p["user"],
        password=p["password"],
        database=p["database"],
        charset="utf8mb4",
        cursorclass=pymysql.cursors.DictCursor,
        autocommit=True,
    )


def ensure_nonconformity_table() -> bool:
    """若表不存在则创建。返回是否成功。"""
    if not _mysql_params()["enabled"]:
        return False
    try:
        import pymysql  # noqa: F401
    except ImportError:
        print("  [DB] 未安装 pymysql，跳过建表。请执行: pip install pymysql")
        return False
    try:
        conn = _connect()
        try:
            with conn.cursor() as cur:
                cur.execute(CREATE_TABLE_SQL)
        finally:
            conn.close()
        return True
    except Exception as exc:
        print(f"  [DB] 建表失败（将跳过写入）: {exc}")
        return False


def insert_nonconformity_records(rows: list[dict[str, Any]]) -> int:
    """
    批量插入不合格记录。rows 每项为字段名字典（与表字段一致，不含 id/created_at/updated_at）。
    返回成功插入条数。
    """
    if not rows:
        return 0
    p = _mysql_params()
    if not p["enabled"]:
        return 0
    try:
        import pymysql  # noqa: F401
    except ImportError:
        print("  [DB] 未安装 pymysql，跳过写入数据库。")
        return 0

    table = p["table"]
    cols = [
        "report_name",
        "report_type",
        "input_full_path",
        "sheet_name",
        "page_no",
        "table_index",
        "sample_no",
        "test_item",
        "standard_text",
        "rule_type",
        "actual_value",
        "fail_reason",
        "excel_row",
        "excel_col",
        "excel_col_letter",
        "xlsx_path",
        "batch_id",
    ]
    placeholders = ", ".join(["%s"] * len(cols))
    sql = f"INSERT INTO `{table}` ({', '.join('`' + c + '`' for c in cols)}) VALUES ({placeholders})"

    tuples = []
    for r in rows:
        tuples.append(tuple(r.get(c, "") for c in cols))

    try:
        conn = _connect()
        try:
            with conn.cursor() as cur:
                cur.executemany(sql, tuples)
        finally:
            conn.close()
        print(f"  [DB] 已写入不合格记录 {len(tuples)} 条 → {p['database']}.{table}")
        return len(tuples)
    except Exception as exc:
        print(f"  [DB] 写入失败: {exc}")
        return 0


def build_record(
    *,
    pdf_path: str,
    report_type: str,
    xlsx_path: str,
    batch_id: str,
    sheet_name: str,
    page_no: int | None,
    table_index: int | None,
    excel_row: int,
    excel_col: int,
    sample_no: str,
    test_item: str,
    standard_text: str,
    rule_type: str,
    actual_value: str,
    fail_reason: str,
    col_letter: str,
) -> dict[str, Any]:
    path = Path(pdf_path)
    return {
        "report_name": path.name,
        "report_type": report_type,
        "input_full_path": str(path.resolve()) if path.exists() else str(path),
        "sheet_name": sheet_name,
        "page_no": page_no,
        "table_index": table_index,
        "sample_no": sample_no[:128] if sample_no else "",
        "test_item": test_item[:768] if test_item else "",
        "standard_text": standard_text[:512] if standard_text else "",
        "rule_type": rule_type[:16] if rule_type else "",
        "actual_value": actual_value[:512] if actual_value else "",
        "fail_reason": fail_reason[:1024] if fail_reason else "",
        "excel_row": excel_row,
        "excel_col": excel_col,
        "excel_col_letter": col_letter,
        "xlsx_path": str(Path(xlsx_path).resolve()) if xlsx_path else "",
        "batch_id": batch_id,
    }
