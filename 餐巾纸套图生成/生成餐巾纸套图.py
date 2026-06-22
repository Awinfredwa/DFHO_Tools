#!/usr/bin/env python3
"""
餐巾纸套图生成工具
================

把每个产品文件夹中的两张图片合成到固定 JPG 模板中。

输入:
  输入/<产品名>/unfolded.jpg  展开图
  输入/<产品名>/folded.jpg    折叠/成品图

如果文件名不标准，但文件夹中恰好有两张图片，会自动用像素更大的作为展开图。
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageOps


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}

INPUT_DIR = "输入"
OUTPUT_DIR = "输出"
TEMPLATE_DIR = "空白模板"


@dataclass(frozen=True)
class Placement:
    source: str
    box: tuple[int, int, int, int]
    opacity: float = 0.92
    texture: float = 0.7


# 坐标来自当前「空白模板」和「F1704完成」的差异区域。
# 每个 box = (left, top, right, bottom)，基于 3072x3072 模板。
SCENES: dict[str, list[Placement]] = {
    "2.jpg": [Placement("folded", (700, 703, 2373, 2370))],
    "3.jpg": [Placement("unfolded", (372, 373, 2703, 2709), opacity=0.90)],
    "4.jpg": [
        Placement("folded", (388, 222, 1257, 1090)),
        Placement("unfolded", (388, 1186, 2079, 2877), opacity=0.90),
    ],
    "5.jpg": [Placement("folded", (768, 781, 2373, 2378))],
    "6.jpg": [Placement("folded", (954, 870, 2385, 2297))],
    "7.jpg": [Placement("folded", (819, 953, 2254, 2386))],
}

DETAIL_CIRCLE_CENTER = (918, 2196)
DETAIL_CIRCLE_RADIUS = 804
DETAIL_BACKGROUND_CROP = (0, 640, 3072, 3072)
DETAIL_CIRCLE_BACKGROUND_COLOR = (224, 229, 236)
DETAIL_TEXTURE_TEMPLATE = "2.jpg"
DETAIL_TEXTURE_SOURCE_BOX = (700, 703, 2373, 2370)
DETAIL_RECT_OPACITY = 0.92
DETAIL_RECT_TEXTURE = 0.7
DETAIL_INSET_RECTS = (
    ((450, 1020, 2010, 2580), -10),
    ((530, 1100, 2090, 2660), 0.0),
)


def print_separator() -> None:
    print("=" * 60)


def get_script_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def find_images(folder: Path) -> list[Path]:
    files = [
        path
        for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    return sorted(files, key=lambda item: item.name.lower())


def classify_product_images(folder: Path) -> tuple[Path, Path]:
    images = find_images(folder)
    if len(images) < 2:
        raise ValueError("需要两张图片：展开图和折叠图")

    unfolded_candidates = [
        p for p in images if any(key in p.stem.lower() for key in ("unfold", "展开", "open", "full"))
    ]
    folded_candidates = [
        p for p in images if any(key in p.stem.lower() for key in ("fold", "折叠", "成品", "photo"))
    ]

    if unfolded_candidates and folded_candidates:
        unfolded = unfolded_candidates[0]
        folded = next((p for p in folded_candidates if p != unfolded), folded_candidates[0])
        if folded != unfolded:
            return unfolded, folded

    # 回退: 两张图时，用像素面积更大的作为展开图。
    sized: list[tuple[int, Path]] = []
    for image_path in images:
        with Image.open(image_path) as image:
            sized.append((image.width * image.height, image_path))
    sized.sort(reverse=True)
    return sized[0][1], sized[1][1]


def resize_to_cover(image: Image.Image, size: tuple[int, int]) -> Image.Image:
    image = ImageOps.exif_transpose(image).convert("RGB")
    src_ratio = image.width / image.height
    dst_ratio = size[0] / size[1]

    if src_ratio > dst_ratio:
        new_height = size[1]
        new_width = round(new_height * src_ratio)
    else:
        new_width = size[0]
        new_height = round(new_width / src_ratio)

    resized = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
    left = (new_width - size[0]) // 2
    top = (new_height - size[1]) // 2
    return resized.crop((left, top, left + size[0], top + size[1]))


def paper_blend(base_crop: Image.Image, art: Image.Image, opacity: float, texture: float) -> Image.Image:
    """Keep napkin paper texture while replacing artwork."""
    base_crop = base_crop.convert("RGB")
    art = ImageEnhance.Contrast(art).enhance(1.04)
    art = ImageEnhance.Color(art).enhance(1.03)
    textured = ImageChops.multiply(art, base_crop)
    textured = Image.blend(art, textured, texture)
    return Image.blend(base_crop, textured, opacity)


def apply_placement(canvas: Image.Image, source_image: Image.Image, placement: Placement) -> None:
    left, top, right, bottom = placement.box
    width = right - left
    height = bottom - top
    art = resize_to_cover(source_image, (width, height))
    base_crop = canvas.crop(placement.box)
    result = paper_blend(base_crop, art, placement.opacity, placement.texture)
    canvas.paste(result, (left, top))


def render_scene(
    template_path: Path,
    output_path: Path,
    unfolded: Image.Image,
    folded: Image.Image,
    placements: Iterable[Placement],
) -> None:
    canvas = Image.open(template_path).convert("RGB")
    sources = {"unfolded": unfolded, "folded": folded}
    for placement in placements:
        apply_placement(canvas, sources[placement.source], placement)
    canvas.save(output_path, quality=95, optimize=True)


def detail_art_tone(image: Image.Image) -> Image.Image:
    image = ImageOps.exif_transpose(image).convert("RGB")
    image = ImageEnhance.Color(image).enhance(0.72)
    image = ImageEnhance.Contrast(image).enhance(0.90)
    image = ImageEnhance.Brightness(image).enhance(1.03)
    return image


def circle_box() -> tuple[int, int, int, int]:
    center_x, center_y = DETAIL_CIRCLE_CENTER
    radius = DETAIL_CIRCLE_RADIUS
    return (
        center_x - radius,
        center_y - radius,
        center_x + radius,
        center_y + radius,
    )


def detail_background(plate_scene: Image.Image) -> Image.Image:
    plate_crop = ImageOps.exif_transpose(plate_scene).convert("RGB").crop(DETAIL_BACKGROUND_CROP)
    return resize_to_cover(plate_crop, (3072, 3072)).convert("RGBA")


def detail_texture_source(texture_template: Image.Image) -> Image.Image:
    return ImageOps.exif_transpose(texture_template).convert("RGB").crop(DETAIL_TEXTURE_SOURCE_BOX)


def textured_detail_art(
    folded: Image.Image,
    texture_source: Image.Image | None,
    size: tuple[int, int],
) -> Image.Image:
    art = resize_to_cover(folded, size)
    if texture_source is None:
        return art.convert("RGBA")

    texture = resize_to_cover(texture_source, size)
    return paper_blend(texture, art, DETAIL_RECT_OPACITY, DETAIL_RECT_TEXTURE).convert("RGBA")


def paste_circle_rects(
    canvas: Image.Image,
    folded: Image.Image,
    texture_source: Image.Image | None,
) -> None:
    circle = circle_box()
    mask = Image.new("L", canvas.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse(circle, fill=255)

    circle_bg = Image.new("RGBA", canvas.size, (*DETAIL_CIRCLE_BACKGROUND_COLOR, 255))
    canvas.paste(circle_bg, (0, 0), mask)

    layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    for box, angle in DETAIL_INSET_RECTS:
        width = box[2] - box[0]
        height = box[3] - box[1]
        art = textured_detail_art(folded, texture_source, (width, height))
        if angle:
            art = art.rotate(angle, resample=Image.Resampling.BICUBIC, expand=True)
        center_x = (box[0] + box[2]) // 2
        center_y = (box[1] + box[3]) // 2
        layer.alpha_composite(art, (center_x - art.width // 2, center_y - art.height // 2))

    layer.putalpha(ImageChops.multiply(layer.getchannel("A"), mask))
    canvas.alpha_composite(layer)


def render_detail(
    output_path: Path,
    folded: Image.Image,
    plate_scene_path: Path,
    texture_template_path: Path | None = None,
) -> None:
    """Create detail image from generated scene 7 instead of a detail template."""
    folded = detail_art_tone(folded)
    texture_source = None
    if texture_template_path is not None and texture_template_path.exists():
        with Image.open(texture_template_path) as texture_template:
            texture_source = detail_texture_source(texture_template)

    with Image.open(plate_scene_path) as plate_scene:
        canvas = detail_background(plate_scene)
        paste_circle_rects(canvas, folded, texture_source)

    draw = ImageDraw.Draw(canvas)
    draw.ellipse(circle_box(), outline=(255, 255, 255), width=16)

    canvas = canvas.convert("RGB")
    canvas.save(output_path, quality=95, optimize=True)


def process_product(product_dir: Path, template_dir: Path, output_root: Path) -> bool:
    print(f"\n📁 正在处理：{product_dir.name}")
    try:
        unfolded_path, folded_path = classify_product_images(product_dir)
    except Exception as exc:
        print(f"  ❌ {exc}")
        return False

    print(f"  展开图：{unfolded_path.name}")
    print(f"  折叠图：{folded_path.name}")

    output_dir = output_root / f"{product_dir.name}完成"
    expected_outputs = set(SCENES.keys()) | {"细节图.jpg"}
    if output_dir.exists():
        existing_outputs = {p.name for p in output_dir.iterdir() if p.is_file()}
        if expected_outputs.issubset(existing_outputs):
            print(f"  ⏭️  已存在完整输出文件夹，跳过：{output_dir.name}")
            return True
        print(f"  🔁 输出文件夹不完整，重新生成：{output_dir.name}")
    output_dir.mkdir(parents=True, exist_ok=True)

    with Image.open(unfolded_path) as unfolded_src, Image.open(folded_path) as folded_src:
        unfolded = ImageOps.exif_transpose(unfolded_src).convert("RGB")
        folded = ImageOps.exif_transpose(folded_src).convert("RGB")

        for file_name, placements in SCENES.items():
            template_path = template_dir / file_name
            if not template_path.exists():
                print(f"  ⚠️  缺少模板：{file_name}，已跳过")
                continue
            render_scene(template_path, output_dir / file_name, unfolded, folded, placements)
            print(f"  ✅ 已生成：{file_name}")

        detail_plate_scene = output_dir / "7.jpg"
        if detail_plate_scene.exists():
            render_detail(
                output_dir / "细节图.jpg",
                folded,
                detail_plate_scene,
                template_dir / DETAIL_TEXTURE_TEMPLATE,
            )
            print("  ✅ 已生成：细节图.jpg")
        else:
            print("  ⚠️  缺少 7.jpg，已跳过：细节图.jpg")

    return True


def main() -> None:
    working_dir = get_script_directory()
    input_dir = working_dir / INPUT_DIR
    output_dir = working_dir / OUTPUT_DIR
    template_dir = working_dir / TEMPLATE_DIR

    print_separator()
    print("餐巾纸套图生成工具")
    print_separator()
    print(f"工作目录：{working_dir}")
    print(f"输入目录：{input_dir}")
    print(f"输出目录：{output_dir}")

    if not template_dir.exists():
        print("❌ 未找到 空白模板 文件夹")
        input("按回车键退出...")
        return

    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    product_dirs = [p for p in input_dir.iterdir() if p.is_dir() and not p.name.startswith(".")]
    if not product_dirs:
        print("❌ 输入文件夹中没有产品子文件夹")
        print("请创建类似：输入/F1704/，并放入展开图和折叠图。")
        input("按回车键退出...")
        return

    print(f"\n找到 {len(product_dirs)} 个产品文件夹：")
    for product_dir in product_dirs:
        print(f"  - {product_dir.name}")

    response = input("\n开始处理？按回车继续，输入 n 取消：").strip().lower()
    if response == "n":
        print("已取消。")
        return

    success_count = 0
    for product_dir in product_dirs:
        if process_product(product_dir, template_dir, output_dir):
            success_count += 1

    print()
    print_separator()
    print("处理完成")
    print_separator()
    print(f"成功处理：{success_count}/{len(product_dirs)}")
    print(f"结果位置：{output_dir}")
    input("按回车键退出...")


if __name__ == "__main__":
    main()
