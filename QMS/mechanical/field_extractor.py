"""
机械性能提取 — 结构化字段提取

直接解析 HTML 表格的 <td> 序列进行匹配，不依赖 Excel 列位置，
避免 OCR 表格 colspan 错位导致的字段映射错误。
"""

from bs4 import BeautifulSoup

from .constants import DATA_ROW_PATTERNS, MECH_FIELD_PATTERNS
from .utils import clean_cell_text, is_skip_value

_SECTION_LABELS = {"机械性能", "mechanical", "力学性能", "力学试验"}


def extract_named_fields(html: str) -> list[dict]:
    """
    从机械性能 HTML 表格中提取结构化字段。

    策略：
      1. 找到主表头行（首个含 ≥2 个力学字段关键词的行），按 <td> 顺序记录字段名
      2. 找到数据行（含「纵向/横向/切向」等关键词），按 <td> 顺序记录值
      3. 前 N 个字段用主表头 TD 索引一对一匹配数据值
      4. 剩余数据值按 MECH_FIELD_PATTERNS 标准顺序依次填充

    返回 [{"方向": "纵向", "屈服强度": "377", "硬度HBW": "164", ...}, ...]
    值为 "—" 表示不参与判断。
    """
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return []
    all_trs = table.find_all("tr")
    if not all_trs:
        return []

    primary_fields: list[str] = []
    primary_found = False
    data_entries: list[tuple[str, list[str]]] = []

    for tr in all_trs:
        tds = tr.find_all(["td", "th"])
        texts = [td.get_text(separator=" ", strip=True) for td in tds]
        row_lower = " ".join(texts).lower()

        direction = _match_direction(row_lower)
        if direction is not None:
            vals = _cells_after_direction(texts)
            data_entries.append((direction, vals))
            continue

        if any(kw in row_lower for kw in ("要求值", "standard")):
            continue

        if primary_found:
            continue

        row_fields = _detect_fields_in_row(texts)
        if len(row_fields) >= 2:
            primary_fields = row_fields
            primary_found = True

    if not data_entries:
        return []

    assigned = set(primary_fields)
    remaining_fields = [f for f, _ in MECH_FIELD_PATTERNS if f not in assigned]

    n_primary = len(primary_fields)
    results: list[dict] = []

    for direction, all_vals in data_entries:
        record: dict[str, str] = {"方向": direction}

        for i, fname in enumerate(primary_fields):
            val = all_vals[i] if i < len(all_vals) else ""
            cleaned = clean_cell_text(val) if val.strip() else ""
            record[fname] = "—" if (not cleaned or is_skip_value(cleaned)) else cleaned

        leftover = all_vals[n_primary:]
        fi = vi = 0
        while fi < len(remaining_fields) and vi < len(leftover):
            val = leftover[vi]
            cleaned = clean_cell_text(val) if val.strip() else ""
            record[remaining_fields[fi]] = "—" if (not cleaned or is_skip_value(cleaned)) else cleaned
            fi += 1
            vi += 1
        while fi < len(remaining_fields):
            record[remaining_fields[fi]] = "—"
            fi += 1

        results.append(record)

    return results


def _match_direction(row_lower: str) -> str | None:
    for d, kws in DATA_ROW_PATTERNS:
        if any(kw in row_lower for kw in kws):
            return d
    return None


def _cells_after_direction(texts: list[str]) -> list[str]:
    """跳过方向单元格，返回剩余值列表。"""
    found = False
    out: list[str] = []
    for t in texts:
        if not found and any(kw in t.lower() for _, kws in DATA_ROW_PATTERNS for kw in kws):
            found = True
            continue
        out.append(t)
    return out


def _detect_fields_in_row(texts: list[str]) -> list[str]:
    """从一行的 <td> 文本中按顺序检测力学字段名。"""
    seen: set[str] = set()
    fields: list[str] = []
    skip_first = True
    for t in texts:
        cl = t.lower().strip()
        if not cl:
            skip_first = False
            continue
        if skip_first and any(lab in cl for lab in _SECTION_LABELS):
            skip_first = False
            continue
        skip_first = False
        for fname, patterns in MECH_FIELD_PATTERNS:
            if fname in seen:
                continue
            if any(pat.lower() in cl for pat in patterns):
                fields.append(fname)
                seen.add(fname)
                break
    return fields
