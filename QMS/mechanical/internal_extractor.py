"""
三方对比 — 内部检测报告字段提取

从内部检测报告的 openpyxl worksheet 中提取全量结构化字段，
包括每列的：标准值（要求值行）和内部检测实测值（试样行）。
"""

import re

from .constants import MECH_FIELD_PATTERNS
from .utils import clean_cell_text, is_skip_value

_SKIP_EXACT = {
    "要求值", "Standard", "standard", "单位", "Unit",
    "Unit Symbol", "实测值", "Actual", "试样编号",
    "Sample No.", "Sample No",
}


def match_to_standard_name(raw: str) -> str | None:
    """
    将 OCR 识别的表头文本（如 '屈服强度 Yield Strength ReL Mpa'）
    映射到 MECH_FIELD_PATTERNS 中的标准字段名（如 '屈服强度'）。
    返回 None 表示未匹配。
    """
    raw_lower = raw.lower()
    for field_name, patterns in MECH_FIELD_PATTERNS:
        for pat in patterns:
            if pat.lower() in raw_lower:
                return field_name
    return None


def _build_row_matrix(ws) -> tuple[list[tuple[int, list[str]]], dict[int, list[str]]]:
    """将 worksheet 转为行列文本矩阵。"""
    rows_data: list[tuple[int, list[str]]] = []
    for r in range(1, ws.max_row + 1):
        vals = []
        for c in range(1, ws.max_column + 1):
            raw = ws.cell(row=r, column=c).value
            vals.append(str(raw).strip() if raw is not None else "")
        rows_data.append((r, vals))
    rows_by_row = {r: vals for r, vals in rows_data}
    return rows_data, rows_by_row


def _infer_field_name(rows_by_row: dict, req_row: int, col_idx: int) -> str:
    """向上追溯同列表头，拼出检测项名称，再尝试规范化为标准字段名。"""
    parts: list[str] = []
    r = req_row - 1
    steps = 0
    while r >= 1 and steps < 12:
        vals = rows_by_row.get(r)
        if vals is None or col_idx >= len(vals):
            r -= 1; steps += 1; continue
        cell = vals[col_idx].strip()
        if not cell or cell in _SKIP_EXACT:
            r -= 1; steps += 1; continue
        parts.append(cell)
        if len(parts) >= 4:
            break
        r -= 1; steps += 1

    raw_header = " ".join(reversed(parts)).strip() if parts else ""

    std = match_to_standard_name(raw_header)
    if std:
        return std
    if raw_header:
        return raw_header[:64]
    return f"列{col_idx + 1}"


def _build_reason(actual: str, rule: dict) -> str:
    """根据规则和实测值生成不合格原因描述。"""
    from .judgment import check_value
    result = check_value(actual, rule)
    if result is False:
        t = rule["type"]
        if t == "range":
            return f"{actual}<{rule['min']}或>{rule['max']}"
        elif t == "gte":
            return f"{actual}<{rule['value']}"
        elif t == "lte":
            return f"{actual}>{rule['value']}"
        elif t == "eq":
            exp = rule["value"]
            exp_s = str(int(exp)) if float(exp).is_integer() else str(exp)
            return f"{actual}≠{exp_s}"
    return ""


def extract_internal_fields(ws) -> list[dict]:
    """
    从内部检测报告的 worksheet 提取全量结构化字段。

    返回列表，每项对应一个检测项：
    {
        "检测项":    "屈服强度",
        "标准值":    "≥355",
        "内部检测值": "410",
        "内部合格":  True | False | None,  # None=无法判定
        "内部不合格原因": ""  # 合格或无法判定时为空串
    }
    """
    from .judgment import parse_requirement, check_value

    rows_data, rows_by_row = _build_row_matrix(ws)

    req_excel_row: int | None = None
    req_vals: list[str] = []
    for excel_row, row_vals in rows_data:
        if any("要求值" in v or "Standard" in v for v in row_vals):
            req_excel_row = excel_row
            req_vals = row_vals
            break

    if req_excel_row is None:
        print("  [内部提取] 未找到「要求值」行，无法提取标准值。")
        return []

    col_rules: dict[int, dict] = {}
    for col_idx, req in enumerate(req_vals):
        rule = parse_requirement(req)
        if rule:
            col_rules[col_idx] = rule

    if not col_rules:
        print("  [内部提取] 要求值行中未解析到有效规则。")
        return []

    actual_rows: list[tuple[int, list[str]]] = []
    for excel_row, row_vals in rows_data:
        if excel_row <= req_excel_row:
            continue
        first_val = next((v for v in row_vals if v), "")
        if re.match(r'^\d{2,}[-–]\d{4,}', first_val):
            actual_rows.append((excel_row, row_vals))

    if not actual_rows:
        print("  [内部提取] 未找到试样数据行（格式：数字-数字，如 22-9637）。")
        return []

    results: list[dict] = []
    seen_fields: set[str] = set()

    for col_idx, rule in col_rules.items():
        field_name = _infer_field_name(rows_by_row, req_excel_row, col_idx)
        standard_text = req_vals[col_idx] if col_idx < len(req_vals) else ""

        for _excel_row, row_vals in actual_rows:
            actual = row_vals[col_idx] if col_idx < len(row_vals) else ""
            actual_clean = clean_cell_text(actual) if actual.strip() else ""

            if is_skip_value(actual_clean):
                actual_clean = "—"

            is_pass: bool | None = None
            reason = ""
            if actual_clean and actual_clean != "—":
                result = check_value(actual_clean, rule)
                is_pass = result
                if result is False:
                    reason = _build_reason(actual_clean, rule)

            dedup_key = field_name
            if dedup_key in seen_fields:
                continue
            seen_fields.add(dedup_key)

            results.append({
                "检测项": field_name,
                "标准值": standard_text,
                "内部检测值": actual_clean,
                "内部合格": is_pass,
                "内部不合格原因": reason,
            })

    return results
