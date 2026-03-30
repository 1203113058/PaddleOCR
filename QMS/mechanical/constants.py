"""
机械性能提取 — 常量定义

所有关键词、字段映射、数据行模式集中管理。
"""

MAX_IMAGE_WIDTH = 1600

MECH_KEYWORDS = {
    "屈服强度", "抗拉强度", "断后伸长率", "断面收缩率",
    "冲击试验温度", "冲击吸收能量", "硬度", "伸长率",
    "机械性能", "力学性能", "力学试验", "机械性能试验",
    "Yield", "Tensile", "Elongation", "Impact",
    "ReL", "Rm", "KV2", "HBW", "HRC", "HB",
    "冲击功", "冲击韧性", "延伸率", "收缩率",
}

MECH_KEYWORDS_LOWER = {k.lower() for k in MECH_KEYWORDS}

MECH_SECTION_START_KEYWORDS = {
    "力学性能", "机械性能", "力学试验", "机械性能试验",
    "mechanical property", "mechanical properties", "mechanical test",
}

MECH_SECTION_END_KEYWORDS = {
    "备注", "结论", "检验结论", "说明", "审核", "批准",
    "remark", "conclusion", "note",
    "金相", "metallic phase", "低倍", "macro-etch", "macro",
    "无损检测", "无损探伤", "超声波", "探伤", "nondestructive",
    "化学成分", "chemical composition",
}

MECH_FIELD_PATTERNS: list[tuple[str, list[str]]] = [
    ("屈服强度",  ["屈服强度", "yield strength", "rel ", "rp0.2"]),
    ("抗拉强度",  ["抗拉强度", "tensile strength", "rm "]),
    ("延伸率",    ["延伸率", "伸长率", "断后伸长率", "elongation"]),
    ("收缩率",    ["收缩率", "断面收缩率", "shrink"]),
    ("冲击值",    ["冲击值", "冲击吸收能量", "冲击吸收", "冲击功", "impact value", "absorbed energy", "akv", "kv2"]),
    ("硬度HBW",   ["hbw"]),
    ("硬度HRC",   ["hrc"]),
]

DATA_ROW_PATTERNS: list[tuple[str, list[str]]] = [
    ("纵向", ["纵向", "longitudinal"]),
    ("切向", ["切向", "tangential"]),
    ("横向", ["横向", "transverse"]),
]
