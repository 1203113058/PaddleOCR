"""
机械性能提取 — 表格筛选与区段提取

从综合质量证明书的大表中定位并切出机械性能区段。
"""

from bs4 import BeautifulSoup

from .constants import (
    MECH_KEYWORDS_LOWER,
    MECH_SECTION_END_KEYWORDS,
    MECH_SECTION_START_KEYWORDS,
)


def html_table_text(html: str) -> str:
    """提取 HTML 表格中全部纯文本。"""
    soup = BeautifulSoup(html, "html.parser")
    return soup.get_text(separator=" ", strip=True)


def is_mechanical_table(html: str, min_keywords: int = 2) -> bool:
    """
    判断一个 HTML 表格是否包含机械性能/力学试验数据。

    综合质量证明书中化学成分和力学性能在同一张大表内，
    因此不排除含化学关键词的表，只要力学关键词命中 ≥ min_keywords 即可。
    """
    text_lower = html_table_text(html).lower()
    hits = 0
    for kw in MECH_KEYWORDS_LOWER:
        if kw in text_lower:
            hits += 1
            if hits >= min_keywords:
                return True
    return False


def extract_mechanical_section_html(html: str) -> str | None:
    """
    从综合大表 HTML 中提取机械性能区段，返回只含力学部分的新 HTML 表格。

    定位策略：
      1. 显式区段标题行（"力学性能"/"机械性能"/"Mechanical Property"等）
      2. 首行含 ≥2 个力学指标关键词（"屈服强度"+"抗拉强度"等）
      3. 向下延伸至结束标记（"金相"/"备注"等）或表尾
    返回 None 表示未能提取（调用方使用完整表格）。
    """
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        return None
    all_rows = table.find_all("tr")
    if not all_rows:
        return None

    start_idx = None
    for i, tr in enumerate(all_rows):
        row_text = tr.get_text(separator=" ", strip=True).lower()
        for header in MECH_SECTION_START_KEYWORDS:
            if header.lower() in row_text:
                start_idx = i
                break
        if start_idx is not None:
            break

    if start_idx is None:
        indicators = {"屈服强度", "抗拉强度", "断后伸长率", "冲击吸收能量",
                      "yield", "tensile", "elongation", "impact", "hardness"}
        for i, tr in enumerate(all_rows):
            row_text = tr.get_text(separator=" ", strip=True).lower()
            if sum(1 for kw in indicators if kw in row_text) >= 2:
                start_idx = i
                break

    if start_idx is None:
        return None

    end_idx = len(all_rows)
    for i in range(start_idx + 2, len(all_rows)):
        row_text = all_rows[i].get_text(separator=" ", strip=True).lower()
        for ending in MECH_SECTION_END_KEYWORDS:
            if ending.lower() in row_text:
                if not any(kw in row_text for kw in MECH_KEYWORDS_LOWER):
                    end_idx = i
                    break
        if end_idx != len(all_rows):
            break

    section_rows = all_rows[start_idx:end_idx]
    if not section_rows:
        return None

    return "<table>" + "".join(str(tr) for tr in section_rows) + "</table>"
