"""
配置管理模块：读写 config.toml，供 pdf2excel.py 和 config_ui.py 共用。

config.toml 与本文件同级，不存在时自动使用内置默认值。
"""

import copy
from pathlib import Path

CONFIG_PATH = Path(__file__).parent / "config.toml"

DEFAULT_CONFIG: dict = {
    "output": {
        "output_dir": "output",
    },
    "ocr": {
        "threshold": 0.991,
        "row_tol": 15,
    },
}


def load_config(config_path: Path | str | None = None) -> dict:
    """读取 config.toml，若不存在或解析失败则返回内置默认值的深拷贝。"""
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    path = Path(config_path) if config_path else CONFIG_PATH

    if not path.exists():
        return cfg

    try:
        try:
            import tomllib  # Python 3.11+
            with open(path, "rb") as f:
                data = tomllib.load(f)
        except ImportError:
            try:
                import tomli  # type: ignore
                with open(path, "rb") as f:
                    data = tomli.load(f)
            except ImportError:
                data = _parse_toml_simple(path)

        if "output" in data:
            cfg["output"].update(data["output"])
        if "ocr" in data:
            cfg["ocr"].update(data["ocr"])

    except Exception as exc:
        print(f"  [警告] 读取 config.toml 失败，使用内置默认值：{exc}")

    return cfg


def save_config(cfg: dict, config_path: Path | str | None = None) -> str:
    """将配置写入 config.toml，返回保存路径字符串。"""
    path = Path(config_path) if config_path else CONFIG_PATH

    output_dir = cfg.get("output", {}).get("output_dir", DEFAULT_CONFIG["output"]["output_dir"])
    threshold  = cfg.get("ocr", {}).get("threshold",  DEFAULT_CONFIG["ocr"]["threshold"])
    row_tol    = cfg.get("ocr", {}).get("row_tol",    DEFAULT_CONFIG["ocr"]["row_tol"])

    content = (
        "# PDF OCR 工具基础配置\n"
        "# 可直接编辑此文件，或通过 config_ui.py 可视化配置页面修改。\n\n"
        "[output]\n"
        f'output_dir = "{output_dir}"\n\n'
        "[ocr]\n"
        f"threshold = {float(threshold)}\n"
        f"row_tol = {int(row_tol)}\n"
    )

    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

    return str(path)


def _parse_toml_simple(path: Path) -> dict:
    """极简 TOML 解析器（仅支持本项目 config.toml 的两节结构）。
    仅作为 tomllib/tomli 均不可用时的最后兜底。
    """
    data: dict = {}
    current_section: str | None = None

    with open(path, encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("[") and line.endswith("]"):
                current_section = line[1:-1].strip()
                data.setdefault(current_section, {})
                continue
            if "=" in line and current_section:
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                try:
                    val = int(val)       # type: ignore[assignment]
                except ValueError:
                    try:
                        val = float(val) # type: ignore[assignment]
                    except ValueError:
                        pass
                data[current_section][key] = val

    return data
