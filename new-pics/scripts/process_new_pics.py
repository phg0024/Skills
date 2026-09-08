#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import io
import math
import sys
import time
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".heic"}


def top_level_images(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_EXTS
    )


def make_contact_sheets(args: argparse.Namespace) -> None:
    src = Path(args.input).expanduser()
    out = Path(args.output).expanduser()
    out.mkdir(parents=True, exist_ok=True)
    files = top_level_images(src)
    font = load_font(26)
    per_page = args.per_page
    cols = args.cols
    thumb_w, thumb_h = args.thumb_width, args.thumb_height
    label_h, margin = 70, 18
    for page in range(math.ceil(len(files) / per_page)):
        batch = files[page * per_page:(page + 1) * per_page]
        rows = math.ceil(len(batch) / cols)
        sheet = Image.new(
            "RGB",
            (cols * (thumb_w + margin) + margin,
             rows * (thumb_h + label_h + margin) + margin),
            "white",
        )
        draw = ImageDraw.Draw(sheet)
        for i, path in enumerate(batch):
            row, col = divmod(i, cols)
            x = margin + col * (thumb_w + margin)
            y = margin + row * (thumb_h + label_h + margin)
            with Image.open(path) as im:
                im = ImageOps.exif_transpose(im).convert("RGB")
                im.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
                bx = x + (thumb_w - im.width) // 2
                by = y + (thumb_h - im.height) // 2
                sheet.paste(im, (bx, by))
            draw.rectangle([x, y, x + thumb_w, y + thumb_h], outline=(200, 200, 200), width=2)
            label = f"{page * per_page + i + 1:02d}  {path.name}"
            draw.text((x + 5, y + thumb_h + 8), label, fill=(0, 0, 0), font=font)
        dst = out / f"contact_{page + 1:02d}.jpg"
        sheet.save(dst, quality=92)
        print(dst)


def load_font(size: int) -> ImageFont.ImageFont:
    for candidate in [
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]:
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            pass
    return ImageFont.load_default()


def apply_rename_map(args: argparse.Namespace) -> None:
    folder = Path(args.input).expanduser()
    mapping = read_mapping(Path(args.map).expanduser())
    missing = [src for src, _ in mapping if not (folder / src).exists()]
    targets = [dst for _, dst in mapping]
    duplicate_targets = sorted({x for x in targets if targets.count(x) > 1})
    existing_targets = [
        dst for src, dst in mapping
        if (folder / dst).exists() and src != dst
    ]
    print(
        f"planned={len(mapping)} missing={len(missing)} "
        f"duplicate_targets={len(duplicate_targets)} existing_targets={len(existing_targets)}"
    )
    if missing:
        print("Missing sources:", *missing, sep="\n  ", file=sys.stderr)
    if duplicate_targets:
        print("Duplicate targets:", *duplicate_targets, sep="\n  ", file=sys.stderr)
    if existing_targets:
        print("Existing targets:", *existing_targets, sep="\n  ", file=sys.stderr)
    if missing or duplicate_targets or existing_targets:
        raise SystemExit(1)
    if args.dry_run:
        for src, dst in mapping:
            print(f"{src} -> {dst}")
        return
    for src, dst in mapping:
        (folder / src).rename(folder / dst)
        print(f"{src} -> {dst}")


def read_mapping(path: Path) -> list[tuple[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None or {"source", "target"} - set(reader.fieldnames):
            raise SystemExit("Rename map must have source,target headers.")
        rows = []
        for row in reader:
            source = (row.get("source") or "").strip()
            target = (row.get("target") or "").strip()
            if source and target:
                rows.append((source, target))
        return rows


def rembg_resize(args: argparse.Namespace) -> None:
    try:
        from rembg import new_session, remove
    except Exception as exc:
        raise SystemExit(f"Could not import rembg. Install/fix rembg first: {exc}") from exc

    src = Path(args.input).expanduser()
    out = Path(args.output).expanduser()
    out.mkdir(parents=True, exist_ok=True)
    files = top_level_images(src)
    session = new_session(args.model, providers=["CPUExecutionProvider"])
    print(f"Processing {len(files)} images into {out}")
    start = time.time()
    for index, path in enumerate(files, 1):
        dst = out / f"{path.stem}.png"
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im).convert("RGB")
            buf = io.BytesIO()
            im.save(buf, format="PNG")
        cut_bytes = remove(buf.getvalue(), session=session)
        with Image.open(io.BytesIO(cut_bytes)) as cut:
            cut = ImageOps.exif_transpose(cut).convert("RGBA")
            cut = resize_longest(cut, args.size)
            cut.save(dst, format="PNG", optimize=True)
        print(f"{index:02d}/{len(files)} {path.name} -> {dst.name} {cut.size[0]}x{cut.size[1]}")
        sys.stdout.flush()
    print(f"Done in {time.time() - start:.1f}s")


def resize_longest(im: Image.Image, size: int) -> Image.Image:
    longest = max(im.size)
    if longest == size:
        return im
    scale = size / longest
    new_size = (max(1, round(im.width * scale)), max(1, round(im.height * scale)))
    return im.resize(new_size, Image.Resampling.LANCZOS)


def convert_to_jpg(args: argparse.Namespace) -> None:
    src = Path(args.input).expanduser()
    out = Path(args.output).expanduser()
    out.mkdir(parents=True, exist_ok=True)
    files = sorted(p for p in src.iterdir() if p.is_file() and p.suffix.lower() == ".png")
    print(f"Converting {len(files)} PNG files to JPG in {out}")
    for path in files:
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im).convert("RGBA")
            bg = Image.new("RGB", im.size, args.background)
            bg.paste(im, mask=im.getchannel("A"))
            dst = out / f"{path.stem}.jpg"
            bg.save(dst, format="JPEG", quality=args.quality, subsampling=0, optimize=True)
        print(f"{path.name} -> {dst.name} {bg.size[0]}x{bg.size[1]}")


def verify(args: argparse.Namespace) -> None:
    folder = Path(args.input).expanduser()
    suffix = ".jpg" if args.format == "jpg" else ".png"
    files = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == suffix)
    bad_dimensions = []
    no_transparency = []
    for path in files:
        with Image.open(path) as im:
            if max(im.size) != args.size:
                bad_dimensions.append((path.name, im.size))
            if args.format == "png":
                if im.mode != "RGBA":
                    no_transparency.append((path.name, im.mode))
                elif im.getextrema()[3] == (255, 255):
                    no_transparency.append((path.name, "opaque alpha"))
    print("files", len(files))
    print("bad_dimensions", bad_dimensions, "count", len(bad_dimensions))
    if args.format == "png":
        print("no_transparency", no_transparency, "count", len(no_transparency))
    if bad_dimensions or no_transparency:
        raise SystemExit(1)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="New product picture workflow helper.")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("contact-sheets")
    p.add_argument("--input", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--per-page", type=int, default=12)
    p.add_argument("--cols", type=int, default=3)
    p.add_argument("--thumb-width", type=int, default=420)
    p.add_argument("--thumb-height", type=int, default=560)
    p.set_defaults(func=make_contact_sheets)

    p = sub.add_parser("apply-rename-map")
    p.add_argument("--input", required=True)
    p.add_argument("--map", required=True)
    p.add_argument("--dry-run", action="store_true")
    p.set_defaults(func=apply_rename_map)

    p = sub.add_parser("rembg-resize")
    p.add_argument("--input", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--size", type=int, default=1920)
    p.add_argument("--model", default="u2net")
    p.set_defaults(func=rembg_resize)

    p = sub.add_parser("jpg")
    p.add_argument("--input", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--quality", type=int, default=95)
    p.add_argument("--background", default="white")
    p.set_defaults(func=convert_to_jpg)

    p = sub.add_parser("verify")
    p.add_argument("--input", required=True)
    p.add_argument("--size", type=int, default=1920)
    p.add_argument("--format", choices=["png", "jpg"], default="jpg")
    p.set_defaults(func=verify)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
