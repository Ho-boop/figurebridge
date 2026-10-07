from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont

import compose_origin_word_figure as core


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render a normalized contact sheet for source-panel style preflight."
    )
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--cols", type=int, required=True)
    parser.add_argument("--files", nargs="*")
    parser.add_argument("--qa-dir", type=Path, required=True)
    parser.add_argument("--preview-width", type=int, default=2600)
    return parser.parse_args()


def resolve_sources(folder: Path, files: list[str] | None) -> list[Path]:
    folder = folder.resolve()
    if files:
        sources = [(folder / name).resolve() for name in files]
    else:
        sources = sorted(
            (path.resolve() for path in folder.iterdir() if path.suffix.lower() in core.SUPPORTED),
            key=core.natural_key,
        )
    if not sources:
        raise RuntimeError("No supported panels were found.")
    for source in sources:
        if not source.is_file():
            raise FileNotFoundError(source)
    kinds = {"emf" if path.suffix.lower() == ".emf" else "raster" for path in sources}
    if len(kinds) != 1:
        raise RuntimeError("Do not mix EMF and raster panels in one preflight.")
    return sources


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def load_font(size: int) -> ImageFont.ImageFont:
    for path in (Path("C:/Windows/Fonts/arialbd.ttf"), Path("C:/Windows/Fonts/arial.ttf")):
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def make_contact_sheet(
    sources: list[Path], previews: list[Path], cols: int, output: Path
) -> None:
    cell_width = 1200
    header_height = 70
    gap = 36
    resized: list[Image.Image] = []
    for preview in previews:
        with Image.open(preview) as image:
            image = image.convert("RGB")
            height = max(1, round(image.height * cell_width / image.width))
            resized.append(image.resize((cell_width, height), Image.Resampling.LANCZOS))
    cell_height = max(image.height for image in resized)
    rows = math.ceil(len(resized) / cols)
    width = gap + cols * (cell_width + gap)
    height = gap + rows * (header_height + cell_height + gap)
    sheet = Image.new("RGB", (width, height), (242, 242, 242))
    draw = ImageDraw.Draw(sheet)
    font = load_font(34)
    for index, (source, image) in enumerate(zip(sources, resized)):
        row, col = divmod(index, cols)
        x = gap + col * (cell_width + gap)
        y = gap + row * (header_height + cell_height + gap)
        draw.text((x, y), source.name, fill="black", font=font)
        image_y = y + header_height
        sheet.paste(image, (x, image_y))
        draw.rectangle(
            (x - 1, image_y - 1, x + image.width, image_y + image.height),
            outline=(150, 150, 150),
            width=2,
        )
    sheet.save(output, "PNG")


def main() -> None:
    args = parse_args()
    if args.cols < 1:
        raise ValueError("--cols must be positive.")
    if args.preview_width < 800:
        raise ValueError("--preview-width must be at least 800.")
    sources = resolve_sources(args.input_dir, args.files)
    qa_dir = args.qa_dir.resolve()
    qa_dir.mkdir(parents=True, exist_ok=True)
    previews = core.render_previews(sources, qa_dir, args.preview_width)
    panels = [core.analyze_panel(source, preview, None) for source, preview in zip(sources, previews)]

    frame_widths = np.array([panel.frame[2] - panel.frame[0] for panel in panels], dtype=float)
    frame_heights = np.array([panel.frame[3] - panel.frame[1] for panel in panels], dtype=float)
    median_width = float(np.median(frame_widths))
    median_height = float(np.median(frame_heights))
    warnings: list[str] = []
    for panel, frame_width, frame_height in zip(panels, frame_widths, frame_heights):
        if panel.touches_source_edge:
            warnings.append(f"{panel.source.name}: visible content touches the source canvas edge")
        if abs(frame_width / median_width - 1) > 0.08:
            warnings.append(f"{panel.source.name}: plot-frame width differs from the group median by more than 8%")
        if abs(frame_height / median_height - 1) > 0.08:
            warnings.append(f"{panel.source.name}: plot-frame height differs from the group median by more than 8%")

    contact_sheet = qa_dir / "preflight-contact-sheet.png"
    make_contact_sheet(sources, previews, args.cols, contact_sheet)
    report = {
        "status": "manual_visual_review_required",
        "contact_sheet": str(contact_sheet),
        "warnings": warnings,
        "manual_checklist": [
            "subfigure-label font/size/weight/offset",
            "axis-title font/size/weight/units",
            "tick-label font/size/weight",
            "legend font/sample-line/marker/spacing/placement",
            "curve/marker/axis/border line widths",
            "canvas and plot-frame proportions",
        ],
        "panels": [
            {
                "source": str(panel.source),
                "sha256": file_hash(panel.source),
                "preview": str(panel.preview),
                "preview_size": [panel.width, panel.height],
                "content_box": list(panel.content_box),
                "plot_frame": list(panel.frame),
            }
            for panel in panels
        ],
    }
    report_path = qa_dir / "preflight-report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(contact_sheet)
    print(report_path)
    print(json.dumps({"warnings": warnings, "status": report["status"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
