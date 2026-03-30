"""
机械性能提取 — 图像预处理（公章去除 + PDF/图片渲染）
"""

from pathlib import Path

import numpy as np
import pypdfium2 as pdfium
from PIL import Image

from .constants import MAX_IMAGE_WIDTH


def remove_red_stamp(img: Image.Image) -> Image.Image:
    """将红色公章区域替换为白色，减少 OCR 干扰。"""
    arr = np.array(img.convert("RGB"), dtype=np.uint8)
    r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
    mask = (r > 150) & (g < 90) & (b < 90)
    arr[mask] = [255, 255, 255]
    removed = int(mask.sum())
    if removed > 0:
        print(f"  [去章] 已清除红色区域 {removed} 个像素")
    return Image.fromarray(arr)


def pdf_to_images(
    pdf_path: str,
    tmp_dir: str | Path,
    remove_stamp: bool = True,
    max_width: int = MAX_IMAGE_WIDTH,
) -> list[tuple[int, str]]:
    """将 PDF 每页渲染为临时图片，返回 [(页码, 图片路径), ...]。"""
    pdf = pdfium.PdfDocument(pdf_path)
    tmp = Path(tmp_dir)
    tmp.mkdir(parents=True, exist_ok=True)
    pages = []
    for i in range(len(pdf)):
        page = pdf[i]
        w, h = page.get_size()
        scale = min(max_width / w, 2.0)
        img = page.render(scale=scale).to_pil()
        if remove_stamp:
            img = remove_red_stamp(img)
        p = tmp / f"_page_{i+1}.png"
        img.save(str(p))
        print(f"  第 {i+1} 页 → {img.size[0]}×{img.size[1]}  准备完成")
        pages.append((i + 1, str(p)))
    return pages


def image_to_pages(
    img_path: str,
    tmp_dir: str | Path,
    remove_stamp: bool = True,
) -> list[tuple[int, str]]:
    """将单张图片预处理后保存，返回 [(1, 图片路径)]。"""
    tmp = Path(tmp_dir)
    tmp.mkdir(parents=True, exist_ok=True)
    dest = tmp / "_page_1.png"
    img = Image.open(img_path).convert("RGB")
    w, h = img.size
    if w > MAX_IMAGE_WIDTH:
        scale = MAX_IMAGE_WIDTH / w
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    if remove_stamp:
        img = remove_red_stamp(img)
    img.save(str(dest))
    print(f"  图片输入 → {img.size[0]}×{img.size[1]}  准备完成")
    return [(1, str(dest))]
