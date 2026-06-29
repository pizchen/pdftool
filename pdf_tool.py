#!/usr/bin/env python3
"""
pdf_tool.py - a single PDF toolkit (N-up + page operations) built on PyMuPDF.

Subcommands:
  nup       Combine multiple pages onto each sheet in an m x n grid (auto-fit).
  merge     Concatenate PDFs in the given order (optional per-file page subset).
  extract   Write a subset of pages to a new PDF.
  delete    Remove a subset of pages; keep the rest.
  reorder   Rearrange pages (swap/move a few, or full --order/--reverse).
  split     Split a PDF into multiple files (per N pages, or by explicit ranges).
  rotate    Rotate selected pages by a multiple of 90 degrees.
  info      Print page count, sizes, rotation and metadata.
  bookmark  Read or edit bookmarks (list/add/delete/update/export/import).

Single dependency (PyMuPDF). Page specs are 1-based, e.g. "1-3,5,8-". The page
operations always write a NEW file; an input is never overwritten.

To use these commands on Office files (PPTX/DOCX/...), first convert them to PDF
with office2pdf.py (a thin LibreOffice wrapper), then run pdf_tool.py on the PDF.

Examples:
    python3 pdf_tool.py nup     -i in.pdf -r 2 -c 2 --orientation landscape --frame
    python3 pdf_tool.py merge   a.pdf b.pdf c.pdf -o all.pdf
    python3 pdf_tool.py merge   cover.pdf:1 body.pdf:2- -o report.pdf
    python3 pdf_tool.py merge   scan1.jpg scan2.png notes.pdf -o bundle.pdf   # images + PDF
    python3 pdf_tool.py extract -i in.pdf -p 1-3,7 -o subset.pdf
    python3 pdf_tool.py delete  -i in.pdf -p 2,4 -o trimmed.pdf
    python3 pdf_tool.py reorder -i in.pdf --swap 3:7 --swap 10:20   # switch a few pages
    python3 pdf_tool.py reorder -i in.pdf --move 500:1              # page 500 to the front
    python3 pdf_tool.py reorder -i in.pdf --move 5-8:end           # move a block to the end
    python3 pdf_tool.py reorder -i in.pdf --order 500,1-499,501-    # explicit permutation
    python3 pdf_tool.py reorder -i in.pdf --reverse
    python3 pdf_tool.py split   -i in.pdf --every 2
    python3 pdf_tool.py split   -i in.pdf --ranges 1-3,4-6,7-
    python3 pdf_tool.py rotate  -i in.pdf --cw                 # whole file, 90 clockwise
    python3 pdf_tool.py rotate  -i in.pdf --ccw -p 2,4-6       # some pages, 90 counter-cw
    python3 pdf_tool.py rotate  -i in.pdf --flip               # 180 (upside down)
    python3 pdf_tool.py rotate  -i in.pdf --angle 270 -p 1     # explicit angle
    python3 pdf_tool.py info    -i in.pdf --per-page
    python3 pdf_tool.py bookmark list   -i in.pdf
    python3 pdf_tool.py bookmark add    -i in.pdf --add 1 1 "Chapter 1" --add 2 2 "Section 1.1"
    python3 pdf_tool.py bookmark delete -i in.pdf --index 2,4-6
    python3 pdf_tool.py bookmark update -i in.pdf --index 3 --title "New title"
    python3 pdf_tool.py bookmark export -i in.pdf -o toc.json
    python3 pdf_tool.py bookmark import -i in.pdf --from toc.json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

try:  # PyMuPDF >= 1.24 prefers the `pymupdf` name; `fitz` still works.
    import pymupdf
except ImportError:  # pragma: no cover
    import fitz as pymupdf

# Points per unit (PDF user-space unit = 1/72 inch).
UNITS = {"pt": 1.0, "mm": 72.0 / 25.4, "cm": 72.0 / 2.54, "in": 72.0}

MM_PER_PT = 25.4 / 72.0


# --------------------------------------------------------------------------- #
# N-up engine (combine multiple source pages onto each output sheet).
# --------------------------------------------------------------------------- #
def parse_page_size(spec: str, unit: str, orientation: str) -> tuple[float, float]:
    """Return (width, height) in points for a named or custom page size.

    A named size (e.g. "A4", "Letter") is looked up via PyMuPDF. A custom size is
    "WxH" optionally suffixed with a unit, e.g. "210x297mm". Orientation forces
    the long/short side assignment.
    """
    spec = spec.strip()
    m = re.fullmatch(r"\s*([\d.]+)\s*[xX*]\s*([\d.]+)\s*(pt|mm|cm|in)?\s*", spec)
    if m:
        w = float(m.group(1))
        h = float(m.group(2))
        u = m.group(3) or unit
        factor = UNITS[u]
        w *= factor
        h *= factor
    else:
        size = pymupdf.paper_size(spec.lower())  # (-1, -1) if unknown
        if size == (-1, -1):
            names = ", ".join(sorted(pymupdf.paper_sizes().keys()))
            raise ValueError(
                f"Unknown page size '{spec}'. Use WxH (e.g. 210x297mm) or one of: {names}"
            )
        w, h = float(size[0]), float(size[1])

    if orientation == "landscape":
        w, h = max(w, h), min(w, h)
    else:  # portrait
        w, h = min(w, h), max(w, h)
    return w, h


def parse_pages(spec: str, total: int) -> list[int]:
    """Parse a 1-based page spec like "1-3,5,8-" into 0-based indices."""
    if not spec or spec.strip().lower() == "all":
        return list(range(total))
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            start = int(a) if a.strip() else 1
            end = int(b) if b.strip() else total
        else:
            start = end = int(part)
        if start < 1 or end < start:
            raise ValueError(f"Invalid page range: '{part}'")
        for p in range(start, min(end, total) + 1):
            out.append(p - 1)
    return out


def fit_rect(cell: "pymupdf.Rect", src_w: float, src_h: float) -> "pymupdf.Rect":
    """Largest rectangle with the src aspect ratio, centered inside `cell`."""
    if src_w <= 0 or src_h <= 0:
        return cell
    scale = min(cell.width / src_w, cell.height / src_h)
    w = src_w * scale
    h = src_h * scale
    x0 = cell.x0 + (cell.width - w) / 2.0
    y0 = cell.y0 + (cell.height - h) / 2.0
    return pymupdf.Rect(x0, y0, x0 + w, y0 + h)


def effective_dims(w: float, h: float, rot: int) -> tuple[float, float]:
    """Width/height of a page after rotating it by `rot` degrees."""
    return (h, w) if rot in (90, 270) else (w, h)


def best_rotation(cell: "pymupdf.Rect", w: float, h: float) -> int:
    """Return 0 or 90 - whichever lets a (w x h) page fit larger in `cell`."""
    s0 = min(cell.width / w, cell.height / h)
    s90 = min(cell.width / h, cell.height / w)
    return 90 if s90 > s0 else 0


def cell_rect(
    row: int,
    col: int,
    page_w: float,
    page_h: float,
    rows: int,
    cols: int,
    margin: float,
    gutter: float,
) -> "pymupdf.Rect":
    """Grid cell rectangle (top-left origin, as used by PyMuPDF)."""
    avail_w = page_w - 2 * margin - (cols - 1) * gutter
    avail_h = page_h - 2 * margin - (rows - 1) * gutter
    cw = avail_w / cols
    ch = avail_h / rows
    x0 = margin + col * (cw + gutter)
    y0 = margin + row * (ch + gutter)
    return pymupdf.Rect(x0, y0, x0 + cw, y0 + ch)


def nup(args: argparse.Namespace) -> None:
    """Combine multiple source pages onto each output sheet (N-up layout)."""
    if args.rows < 1 or args.cols < 1:
        raise ValueError("rows and cols must be >= 1")

    unit_factor = UNITS[args.unit]
    margin = args.margin * unit_factor
    gutter = args.gutter * unit_factor

    page_w, page_h = parse_page_size(args.page_size, args.unit, args.orientation)

    if page_w - 2 * margin - (args.cols - 1) * gutter <= 0 or \
       page_h - 2 * margin - (args.rows - 1) * gutter <= 0:
        raise ValueError("Margins/gutter leave no room for the grid; reduce them.")

    src = pymupdf.open(args.input)
    try:
        if not src.is_pdf:
            raise ValueError("Input is not a PDF file.")
        page_indices = parse_pages(args.pages, src.page_count)
        if not page_indices:
            raise ValueError("No pages selected.")

        per_sheet = args.rows * args.cols
        out = pymupdf.open()
        try:
            for start in range(0, len(page_indices), per_sheet):
                chunk = page_indices[start:start + per_sheet]
                sheet = out.new_page(width=page_w, height=page_h)
                for slot, src_idx in enumerate(chunk):
                    if args.order == "col":
                        col = slot // args.rows
                        row = slot % args.rows
                    else:  # row-major
                        row = slot // args.cols
                        col = slot % args.cols

                    cell = cell_rect(
                        row, col, page_w, page_h,
                        args.rows, args.cols, margin, gutter,
                    )
                    src_page = src[src_idx]
                    sr = src_page.rect  # visible rect (accounts for rotation)
                    if args.rotate == "auto":
                        rot = best_rotation(cell, sr.width, sr.height)
                    else:
                        rot = int(args.rotate)
                    eff_w, eff_h = effective_dims(sr.width, sr.height, rot)
                    target = fit_rect(cell, eff_w, eff_h)
                    sheet.show_pdf_page(
                        target, src, src_idx,
                        rotate=rot,
                        keep_proportion=True,
                    )
                    if args.frame:
                        sheet.draw_rect(
                            target,
                            color=(0, 0, 0),
                            width=args.frame_width,
                        )

            out.save(args.output, garbage=4, deflate=True)
        finally:
            out.close()
    finally:
        src.close()

    sheets = (len(page_indices) + per_sheet - 1) // per_sheet
    extras = []
    if args.rotate != "0":
        extras.append(f"rotate={args.rotate}")
    if args.frame:
        extras.append("framed")
    extra = (", " + ", ".join(extras)) if extras else ""
    print(
        f"Done: {len(page_indices)} source page(s) -> {sheets} sheet(s) "
        f"({args.rows}x{args.cols} = {per_sheet}-up, {args.orientation} "
        f"{args.page_size}{extra}) -> {args.output}"
    )


def open_pdf(path: str) -> "pymupdf.Document":
    """Open a path as a PDF or raise a clear error."""
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    doc = pymupdf.open(path)
    if not doc.is_pdf:
        doc.close()
        raise ValueError(f"Not a PDF file: {path}")
    return doc


def ensure_not_input(output: str, *inputs: str) -> None:
    """Refuse to write an output on top of any input file."""
    out_abs = os.path.abspath(output)
    for inp in inputs:
        if os.path.abspath(inp) == out_abs:
            raise ValueError(
                f"Output '{output}' would overwrite an input file; choose a "
                f"different -o/--output."
            )


def save_new(doc: "pymupdf.Document", output: str) -> None:
    """Save to a (new) path, creating parent directories as needed."""
    parent = os.path.dirname(os.path.abspath(output))
    os.makedirs(parent, exist_ok=True)
    doc.save(output, garbage=4, deflate=True)


def parse_merge_item(token: str) -> tuple[str, str | None]:
    """Split a merge token 'FILE[:pagespec]' into (path, pagespec|None).

    Splits on the LAST colon only when the right side looks like a page spec
    (digits/commas/dashes), so Windows paths like 'C:\\a.pdf' stay intact.
    """
    if ":" in token:
        left, right = token.rsplit(":", 1)
        if left and re.fullmatch(r"[\d\s,\-]+", right) and any(c.isdigit() for c in right):
            return left, right
    return token, None


# Image formats `merge` accepts. PyMuPDF natively decodes JPEG/PNG/GIF/BMP/TIFF/
# JPEG2000/PNM; WEBP/HEIC and similar fall back to Pillow when it is installed.
IMAGE_EXTS = {
    ".jpg", ".jpeg", ".jpe", ".jfif",
    ".png", ".gif", ".bmp", ".dib",
    ".tif", ".tiff", ".pnm", ".pgm", ".pbm", ".ppm", ".pam",
    ".jp2", ".jpx", ".j2k", ".jpf", ".jpc",
    ".webp", ".heic", ".heif",
}


def _pillow_to_png(path: str) -> "tuple[int, int, bytes] | None":
    """Decode an image with Pillow (optional) -> (width, height, png_bytes).

    Fallback only for formats PyMuPDF cannot open natively (e.g. WEBP, HEIC).
    Returns None if Pillow is unavailable or the file cannot be decoded.
    """
    try:
        from PIL import Image
    except ImportError:
        return None
    try:  # enable HEIC/HEIF if the optional plug-in is present
        from pillow_heif import register_heif_opener
        register_heif_opener()
    except Exception:
        pass
    try:
        import io
        with Image.open(path) as im:
            im.load()
            has_alpha = im.mode in ("RGBA", "LA", "PA") or (
                im.mode == "P" and "transparency" in im.info
            )
            im = im.convert("RGBA" if has_alpha else "RGB")
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            return im.width, im.height, buf.getvalue()
    except Exception:
        return None


def _load_image(path: str) -> "tuple[float, float, dict]":
    """Return (width, height, insert_kwargs) for an image input.

    Prefers PyMuPDF's native decoders (which keep the original compressed image
    stream), then falls back to Pillow for other formats.
    """
    try:
        with pymupdf.open(path) as doc:
            if doc.is_pdf:
                raise ValueError(f"'{path}' is a PDF, not an image.")
            rect = doc[0].rect
            return rect.width, rect.height, {"filename": path}
    except ValueError:
        raise
    except Exception:
        pass  # not natively decodable; try Pillow below
    decoded = _pillow_to_png(path)
    if decoded is None:
        raise ValueError(
            f"Could not decode image '{path}'. PyMuPDF handles JPEG/PNG/GIF/BMP/"
            f"TIFF/JPEG2000/PNM directly; for WEBP/HEIC and similar install Pillow "
            f"(`pip install pillow`, plus `pillow-heif` for HEIC)."
        )
    w, h, png = decoded
    return float(w), float(h), {"stream": png}


def add_image_page(
    out: "pymupdf.Document",
    path: str,
    page_size: str,
    orientation: str,
    unit: str,
    margin_units: float,
) -> None:
    """Append one page (of the target size) holding the auto-fitted, centered image."""
    iw, ih, insert_kwargs = _load_image(path)
    orient = ("landscape" if iw >= ih else "portrait") if orientation == "auto" else orientation
    page_w, page_h = parse_page_size(page_size, unit, orient)
    margin = margin_units * UNITS[unit]
    if page_w - 2 * margin <= 0 or page_h - 2 * margin <= 0:
        raise ValueError("Margin too large for the image page size.")
    page = out.new_page(width=page_w, height=page_h)
    content = pymupdf.Rect(margin, margin, page_w - margin, page_h - margin)
    # keep_proportion=True fits the image inside `content` preserving aspect ratio
    # and centers it (verified behavior).
    page.insert_image(content, keep_proportion=True, **insert_kwargs)


def cmd_merge(args: argparse.Namespace) -> None:
    items = [parse_merge_item(t) for t in args.inputs]
    ensure_not_input(args.output, *[p for p, _ in items])
    out = pymupdf.open()
    try:
        total = n_pdf = n_img = 0
        for path, spec in items:
            if not os.path.exists(path):
                raise FileNotFoundError(path)
            if os.path.splitext(path)[1].lower() in IMAGE_EXTS:
                if spec:
                    raise ValueError(
                        f"A page spec (':{spec}') is not valid for image input '{path}'."
                    )
                add_image_page(out, path, args.page_size, args.orientation,
                               args.unit, args.margin)
                total += 1
                n_img += 1
            else:
                src = open_pdf(path)
                try:
                    if spec:
                        idx = parse_pages(spec, src.page_count)
                        if not idx:
                            raise ValueError(f"No pages selected from '{path}' (spec '{spec}').")
                        src.select(idx)
                    out.insert_pdf(src)
                    total += src.page_count
                    n_pdf += 1
                finally:
                    src.close()
        if total == 0:
            raise ValueError("No pages to merge.")
        save_new(out, args.output)
    finally:
        out.close()
    kinds = []
    if n_pdf:
        kinds.append(f"{n_pdf} PDF")
    if n_img:
        kinds.append(f"{n_img} image")
    detail = f" ({', '.join(kinds)})" if kinds else ""
    print(f"Merged {len(items)} input(s){detail} -> {total} page(s) -> {args.output}")


def cmd_extract(args: argparse.Namespace) -> None:
    ensure_not_input(args.output, args.input)
    doc = open_pdf(args.input)
    try:
        idx = parse_pages(args.pages, doc.page_count)
        if not idx:
            raise ValueError("No pages selected.")
        doc.select(idx)
        save_new(doc, args.output)
        n = doc.page_count
    finally:
        doc.close()
    print(f"Extracted {n} page(s) ('{args.pages}') -> {args.output}")


def cmd_delete(args: argparse.Namespace) -> None:
    ensure_not_input(args.output, args.input)
    doc = open_pdf(args.input)
    try:
        total = doc.page_count
        remove = set(parse_pages(args.pages, total))
        keep = [i for i in range(total) if i not in remove]
        if not keep:
            raise ValueError("Refusing to delete every page (nothing would remain).")
        doc.select(keep)
        save_new(doc, args.output)
        n = doc.page_count
    finally:
        doc.close()
    print(f"Deleted {len(remove)} page(s), kept {n} -> {args.output}")


def parse_swaps(values: list[str], total: int) -> list[tuple[int, int]]:
    """Parse repeatable/comma-separated 'A:B' swap pairs (1-based page numbers)."""
    pairs: list[tuple[int, int]] = []
    for value in values:
        for item in value.split(","):
            item = item.strip()
            if not item:
                continue
            if ":" not in item:
                raise ValueError(f"Bad --swap '{item}'; expected 'A:B'.")
            a_str, _, b_str = item.partition(":")
            try:
                a, b = int(a_str), int(b_str)
            except ValueError:
                raise ValueError(f"Bad --swap '{item}'; A and B must be integers.")
            if not (1 <= a <= total) or not (1 <= b <= total):
                raise ValueError(f"--swap '{item}' out of range 1..{total}.")
            if a == b:
                raise ValueError(f"--swap '{item}' swaps a page with itself.")
            pairs.append((a, b))
    return pairs


def parse_moves(values: list[str], total: int) -> list[tuple[int, int, object]]:
    """Parse 'SRC:DEST' moves into (start, end, dest) with 1-based inclusive SRC.

    SRC is a page ('50') or contiguous block ('5-8'); DEST is a 1-based target
    position, or the keyword 'start' / 'end'.
    """
    moves: list[tuple[int, int, object]] = []
    for value in values:
        for item in value.split(","):
            item = item.strip()
            if not item:
                continue
            if ":" not in item:
                raise ValueError(f"Bad --move '{item}'; expected 'SRC:DEST'.")
            src, _, dest = item.partition(":")
            src, dest = src.strip(), dest.strip().lower()
            try:
                if "-" in src:
                    a_str, _, b_str = src.partition("-")
                    start, end = int(a_str), int(b_str)
                else:
                    start = end = int(src)
            except ValueError:
                raise ValueError(f"Bad --move '{item}'; SRC must be a page or 'a-b' block.")
            if not (1 <= start <= end <= total):
                raise ValueError(f"--move '{item}' SRC out of range 1..{total}.")
            target: object
            if dest in ("start", "end"):
                target = dest
            else:
                try:
                    target = int(dest)
                except ValueError:
                    raise ValueError(f"Bad --move '{item}'; DEST must be an int, 'start', or 'end'.")
            moves.append((start, end, target))
    return moves


def apply_swaps(order: list[int], swaps: list[tuple[int, int]]) -> list[int]:
    """Apply swaps simultaneously by original page number; pairs must be disjoint."""
    seen: set[int] = set()
    for a, b in swaps:
        if a in seen or b in seen:
            raise ValueError(
                f"--swap pages must be disjoint; page {a if a in seen else b} appears twice."
            )
        seen.update((a, b))
    result = list(order)
    for a, b in swaps:
        result[a - 1], result[b - 1] = result[b - 1], result[a - 1]
    return result


def apply_moves(order: list[int], moves: list[tuple[int, int, object]]) -> list[int]:
    """Apply moves sequentially; SRC/DEST refer to the current working positions.

    A moved block ends up starting exactly at 1-based position DEST in the
    resulting list (DEST is interpreted after the block is removed).
    """
    work = list(order)
    for start, end, dest in moves:
        n = len(work)
        if not (1 <= start <= end <= n):
            raise ValueError(f"--move SRC {start}-{end} out of range 1..{n}.")
        block = work[start - 1:end]
        del work[start - 1:end]
        if dest == "start":
            idx = 0
        elif dest == "end":
            idx = len(work)
        else:
            d = int(dest)  # type: ignore[arg-type]
            if not (1 <= d <= len(work) + 1):
                raise ValueError(
                    f"--move DEST {d} out of range 1..{len(work) + 1} "
                    f"(after removing the block); use 'end' for the last position."
                )
            idx = d - 1
        work[idx:idx] = block
    return work


def cmd_reorder(args: argparse.Namespace) -> None:
    ensure_not_input(args.output, args.input)
    doc = open_pdf(args.input)
    try:
        total = doc.page_count
        edits = bool(args.swap) or bool(args.move)
        chosen = [bool(args.reverse), bool(args.order), edits]
        if sum(chosen) == 0:
            raise ValueError("Specify one of --reverse, --order, --swap, or --move.")
        if sum(chosen) > 1:
            raise ValueError(
                "--reverse, --order, and --swap/--move are different modes; use one "
                "per run (swaps and moves may be combined with each other)."
            )

        if args.reverse:
            order = list(range(total - 1, -1, -1))
            how = "reverse"
        elif args.order:
            order = parse_pages(args.order, total)
            if not args.allow_partial and sorted(order) != list(range(total)):
                raise ValueError(
                    f"--order must be a permutation of all {total} page(s); got "
                    f"{len(order)} index/indices. Use --allow-partial to drop or "
                    f"duplicate pages on purpose."
                )
            how = f"order '{args.order}'"
        else:
            order = list(range(total))  # identity; only the deltas below change it
            parts = []
            if args.swap:
                swaps = parse_swaps(args.swap, total)
                order = apply_swaps(order, swaps)
                parts.append(f"{len(swaps)} swap(s)")
            if args.move:
                moves = parse_moves(args.move, total)
                order = apply_moves(order, moves)
                parts.append(f"{len(moves)} move(s)")
            how = ", ".join(parts)

        if not order:
            raise ValueError("Empty page order.")
        doc.select(order)
        save_new(doc, args.output)
        n = doc.page_count
    finally:
        doc.close()
    print(f"Reordered ({how}) -> {n} page(s) -> {args.output}")


def cmd_split(args: argparse.Namespace) -> None:
    doc = open_pdf(args.input)
    try:
        total = doc.page_count
        base = os.path.splitext(os.path.basename(args.input))[0]
        outdir = args.outdir or os.path.dirname(os.path.abspath(args.input))
        os.makedirs(outdir, exist_ok=True)

        groups: list[list[int]] = []
        if args.ranges:
            for part in args.ranges.split(","):
                part = part.strip()
                if part:
                    idx = parse_pages(part, total)
                    if idx:
                        groups.append(idx)
        else:
            if args.every < 1:
                raise ValueError("--every must be >= 1.")
            for start in range(0, total, args.every):
                groups.append(list(range(start, min(start + args.every, total))))

        if not groups:
            raise ValueError("Nothing to split.")

        written: list[str] = []
        for grp in groups:
            sub = pymupdf.open()
            try:
                for i in grp:
                    sub.insert_pdf(doc, from_page=i, to_page=i)
                name = f"{base}_{grp[0] + 1}-{grp[-1] + 1}.pdf"
                path = os.path.join(outdir, name)
                ensure_not_input(path, args.input)
                save_new(sub, path)
                written.append(path)
            finally:
                sub.close()
    finally:
        doc.close()
    print(f"Split '{args.input}' into {len(written)} file(s) in {outdir}:")
    for p in written:
        print(f"  {p}")


def cmd_rotate(args: argparse.Namespace) -> None:
    ensure_not_input(args.output, args.input)
    angle = args.angle if args.angle is not None else args.shortcut
    if angle % 90 != 0:
        raise ValueError("Rotation angle must be a multiple of 90 (e.g. 90, 180, 270, -90).")
    doc = open_pdf(args.input)
    try:
        idx = parse_pages(args.pages, doc.page_count)
        if not idx:
            raise ValueError("No pages selected.")
        for i in idx:
            page = doc[i]
            new_rot = (angle % 360) if args.absolute else (page.rotation + angle) % 360
            page.set_rotation(new_rot)
        save_new(doc, args.output)
    finally:
        doc.close()
    scope = "whole file" if args.pages.strip().lower() == "all" else f"pages '{args.pages}'"
    verb = "set to" if args.absolute else "by"
    print(f"Rotated {len(idx)} page(s) ({scope}) {verb} {angle % 360} deg -> {args.output}")


def cmd_info(args: argparse.Namespace) -> None:
    doc = open_pdf(args.input)
    try:
        size = os.path.getsize(args.input)
        print(f"File:  {args.input}")
        print(f"Size:  {size / 1024:.1f} KiB ({size} bytes)")
        print(f"Pages: {doc.page_count}")
        meta = doc.metadata or {}
        shown = [(k, meta.get(k)) for k in ("title", "author", "subject", "creator", "producer")]
        for key, val in shown:
            if val:
                print(f"  {key}: {val}")
        if args.per_page:
            print("Pages (1-based):")
            for i in range(doc.page_count):
                page = doc[i]
                r = page.rect
                print(
                    f"  {i + 1:>4}: {r.width:7.1f}x{r.height:7.1f} pt  "
                    f"({r.width * MM_PER_PT:6.1f}x{r.height * MM_PER_PT:6.1f} mm)  "
                    f"rot={page.rotation}"
                )
    finally:
        doc.close()


# --------------------------------------------------------------------------- #
# Bookmarks (PDF outline / table of contents).
#
# PyMuPDF stores the outline as a flat list of [level, title, page] rows
# (1-based pages; page -1 means "no destination"). The 1-based `level` encodes
# the tree: the first row must be level 1 and each row may deepen the level by
# at most 1. Every editing action writes a NEW pdf; the input is never modified.
# --------------------------------------------------------------------------- #
def get_bookmarks(doc: "pymupdf.Document") -> list:
    """Return the document outline as a list of [level, title, page] rows."""
    return doc.get_toc(simple=True)


def validate_bookmarks(toc: list, total: int) -> None:
    """Validate a TOC before applying it, raising clear 1-based row errors.

    Checks positive integer levels, a legal hierarchy (the first row is level 1
    and each level steps up by at most 1) and page numbers in 1..total (or -1 for
    "no destination"). Without this, set_toc would clamp out-of-range pages
    silently and give only a terse hierarchy message.
    """
    prev = 0
    for i, entry in enumerate(toc):
        row = i + 1
        level, title, page = entry[0], entry[1], entry[2]
        if not isinstance(level, int) or level < 1:
            raise ValueError(f"bookmark row {row}: level must be an integer >= 1 (got {level!r}).")
        if level > prev + 1:
            raise ValueError(
                f"bookmark row {row} ('{title}'): level {level} jumps from {prev}; the "
                f"first bookmark must be level 1 and each level may only increase by 1."
            )
        if not isinstance(page, int):
            raise ValueError(f"bookmark row {row} ('{title}'): page must be an integer (got {page!r}).")
        if page != -1 and not (1 <= page <= total):
            raise ValueError(
                f"bookmark row {row} ('{title}'): page {page} is out of range 1..{total} "
                f"(use -1 for a bookmark with no page destination)."
            )
        prev = level


def parse_index_spec(spec: str, n: int) -> list:
    """Parse a 1-based bookmark-index spec ('3', '2,4-6', 'all') into 0-based indices."""
    return parse_pages(spec, n)


def write_toc(doc: "pymupdf.Document", toc: list, output: str) -> None:
    """Validate `toc`, set it as the outline, and save to `output` (a new file)."""
    validate_bookmarks(toc, doc.page_count)
    doc.set_toc(toc)
    save_new(doc, output)


def default_bookmark_output(args: argparse.Namespace) -> str:
    """Default a missing -o for bookmark write actions to <input>_bookmarks.pdf."""
    if not getattr(args, "output", None):
        base, _ = os.path.splitext(args.input)
        args.output = f"{base}_bookmarks.pdf"
    return args.output


def _bookmark_loc(page: object) -> str:
    """Human-readable label for a bookmark's page destination."""
    return f"p.{page}" if isinstance(page, int) and page > 0 else "no dest"


def cmd_bm_list(args: argparse.Namespace) -> None:
    doc = open_pdf(args.input)
    try:
        toc = get_bookmarks(doc)
        if args.json:
            data = [
                {"index": i + 1, "level": lvl, "title": title, "page": page}
                for i, (lvl, title, page) in enumerate(toc)
            ]
            print(json.dumps(data, ensure_ascii=False, indent=2))
            return
        if not toc:
            print("(no bookmarks)")
            return
        width = len(str(len(toc)))
        for i, (lvl, title, page) in enumerate(toc):
            indent = "  " * (lvl - 1)
            print(f"{i + 1:>{width}}: {indent}{title}  ({_bookmark_loc(page)})")
    finally:
        doc.close()


def _jsonify_dest(dest: dict) -> dict:
    """Make a detailed get_toc dest dict JSON-serializable (Point -> [x, y])."""
    out: dict = {}
    for key, val in dest.items():
        if hasattr(val, "x") and hasattr(val, "y"):  # pymupdf.Point
            out[key] = [val.x, val.y]
        else:
            out[key] = val
    return out


def cmd_bm_export(args: argparse.Namespace) -> None:
    doc = open_pdf(args.input)
    try:
        if args.details:
            data = []
            for entry in doc.get_toc(simple=False):
                lvl, title, page = entry[0], entry[1], entry[2]
                dest = entry[3] if len(entry) > 3 else {}
                data.append({"level": lvl, "title": title, "page": page,
                             "dest": _jsonify_dest(dest)})
        else:
            data = [{"level": lvl, "title": title, "page": page}
                    for lvl, title, page in get_bookmarks(doc)]
    finally:
        doc.close()
    text = json.dumps(data, ensure_ascii=False, indent=2)
    out = args.output
    if not out:
        base, _ = os.path.splitext(args.input)
        out = f"{base}_bookmarks.json"
    if out == "-":
        print(text)
        return
    ensure_not_input(out, args.input)
    parent = os.path.dirname(os.path.abspath(out))
    os.makedirs(parent, exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(text + "\n")
    print(f"Exported {len(data)} bookmark(s) -> {out}")


def _insert_sorted_by_page(toc: list, entry: list) -> None:
    """Insert `entry` into `toc` (in place) keeping it ordered by page (stable).

    No-destination entries (page -1) sort to the end. The new entry goes before
    the first existing row whose page is strictly greater.
    """
    page = entry[2]
    key = page if (isinstance(page, int) and page > 0) else float("inf")
    insert_at = len(toc)
    for i, existing in enumerate(toc):
        ep = existing[2]
        ekey = ep if (isinstance(ep, int) and ep > 0) else float("inf")
        if ekey > key:
            insert_at = i
            break
    toc.insert(insert_at, entry)


def cmd_bm_add(args: argparse.Namespace) -> None:
    if not args.add:
        raise ValueError('Provide at least one --add PAGE LEVEL "TITLE".')
    default_bookmark_output(args)
    ensure_not_input(args.output, args.input)
    doc = open_pdf(args.input)
    try:
        total = doc.page_count
        toc = [list(e) for e in get_bookmarks(doc)]
        new_entries = []
        for page_s, level_s, title in args.add:
            try:
                page, level = int(page_s), int(level_s)
            except ValueError:
                raise ValueError(
                    f'--add expects PAGE LEVEL "TITLE" with integer PAGE and LEVEL; '
                    f"got PAGE={page_s!r} LEVEL={level_s!r}."
                )
            if page != -1 and not (1 <= page <= total):
                raise ValueError(f"--add page {page} out of range 1..{total} (use -1 for no destination).")
            if level < 1:
                raise ValueError(f"--add level {level} must be >= 1.")
            new_entries.append([level, title, page])

        if args.at is not None:
            pos = args.at - 1
            if not (0 <= pos <= len(toc)):
                raise ValueError(f"--at {args.at} out of range 1..{len(toc) + 1}.")
            toc[pos:pos] = new_entries
        else:
            for entry in new_entries:
                _insert_sorted_by_page(toc, entry)

        write_toc(doc, toc, args.output)
        added, final_count = len(new_entries), len(toc)
    finally:
        doc.close()
    print(f"Added {added} bookmark(s); {final_count} total -> {args.output}")


def _delete_keep_children(toc: list, sel: set) -> list:
    """Remove selected rows but keep descendants, promoting them to stay valid.

    Each kept row's level is reduced by the number of its deleted ancestors, so
    orphaned children move up to fill the gap left by a removed parent.
    """
    result = []
    stack: list = []  # ancestor path as (level, is_deleted)
    for i, entry in enumerate(toc):
        level = entry[0]
        while stack and stack[-1][0] >= level:
            stack.pop()
        deleted_ancestors = sum(1 for lvl, dele in stack if dele)
        is_deleted = i in sel
        stack.append((level, is_deleted))
        if not is_deleted:
            new_level = max(1, level - deleted_ancestors)
            result.append([new_level, entry[1], entry[2]])
    return result


def cmd_bm_delete(args: argparse.Namespace) -> None:
    default_bookmark_output(args)
    ensure_not_input(args.output, args.input)
    doc = open_pdf(args.input)
    try:
        toc = [list(e) for e in get_bookmarks(doc)]
        if args.all:
            new_toc: list = []
        else:
            if not args.index:
                raise ValueError("Specify --index SPEC (e.g. '2,4-6') or --all.")
            if not toc:
                raise ValueError("Document has no bookmarks to delete.")
            sel = set(parse_index_spec(args.index, len(toc)))
            if not sel:
                raise ValueError("No bookmarks selected by --index.")
            if args.keep_children:
                new_toc = _delete_keep_children(toc, sel)
            else:
                remove = set(sel)
                for i in sel:
                    base = toc[i][0]
                    j = i + 1
                    while j < len(toc) and toc[j][0] > base:
                        remove.add(j)
                        j += 1
                new_toc = [list(e) for k, e in enumerate(toc) if k not in remove]
        removed = len(toc) - len(new_toc)
        if removed == 0:
            raise ValueError("No bookmarks were removed.")
        write_toc(doc, new_toc, args.output)
        remaining = len(new_toc)
    finally:
        doc.close()
    print(f"Deleted {removed} bookmark(s); {remaining} remaining -> {args.output}")


def cmd_bm_update(args: argparse.Namespace) -> None:
    if args.title is None and args.page is None and args.level is None:
        raise ValueError("Specify at least one of --title, --page, --level.")
    default_bookmark_output(args)
    ensure_not_input(args.output, args.input)
    doc = open_pdf(args.input)
    try:
        toc = [list(e) for e in get_bookmarks(doc)]
        if not toc:
            raise ValueError("Document has no bookmarks to update.")
        sel = parse_index_spec(args.index, len(toc))
        if not sel:
            raise ValueError("No bookmarks selected by --index.")
        for i in sel:
            if args.title is not None:
                toc[i][1] = args.title
            if args.page is not None:
                toc[i][2] = args.page
            if args.level is not None:
                toc[i][0] = args.level
        write_toc(doc, toc, args.output)
        n = len(sel)
    finally:
        doc.close()
    print(f"Updated {n} bookmark(s) -> {args.output}")


def _toc_row_from_json(item: object, n: int) -> list:
    """Build a [level, title, page(, dest)] TOC row from one JSON entry."""
    if isinstance(item, dict):
        for key in ("level", "title", "page"):
            if key not in item:
                raise ValueError(f"bookmark entry {n}: missing '{key}'.")
        level, title, page, dest = item["level"], item["title"], item["page"], item.get("dest")
    elif isinstance(item, (list, tuple)):
        if len(item) < 3:
            raise ValueError(f"bookmark entry {n}: array must be [level, title, page].")
        level, title, page = item[0], item[1], item[2]
        dest = item[3] if len(item) > 3 else None
    else:
        raise ValueError(f"bookmark entry {n}: must be a JSON object or array.")
    if not isinstance(level, int) or not isinstance(page, int):
        raise ValueError(f"bookmark entry {n}: 'level' and 'page' must be integers.")
    title = str(title)
    if isinstance(dest, dict):
        clean = dict(dest)
        to = clean.get("to")
        if isinstance(to, (list, tuple)) and len(to) == 2:
            clean["to"] = pymupdf.Point(float(to[0]), float(to[1]))
        return [level, title, page, clean]
    return [level, title, page]


def cmd_bm_import(args: argparse.Namespace) -> None:
    default_bookmark_output(args)
    ensure_not_input(args.output, args.input)
    if args.source == "-":
        text = sys.stdin.read()
    else:
        if not os.path.exists(args.source):
            raise FileNotFoundError(args.source)
        with open(args.source, encoding="utf-8") as fh:
            text = fh.read()
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON: {exc}")
    if not isinstance(data, list):
        raise ValueError("Bookmark JSON must be a list of entries.")
    toc = [_toc_row_from_json(item, i + 1) for i, item in enumerate(data)]
    doc = open_pdf(args.input)
    try:
        write_toc(doc, toc, args.output)
        count = len(toc)
    finally:
        doc.close()
    src = "stdin" if args.source == "-" else f"'{args.source}'"
    print(f"Imported {count} bookmark(s) from {src} -> {args.output}")


def cmd_nup(args: argparse.Namespace) -> None:
    nup(args)


def add_nup_arguments(p: argparse.ArgumentParser) -> None:
    """Add the N-up options to the `nup` subcommand's parser."""
    p.add_argument("-i", "--input", required=True, help="Input PDF path.")
    p.add_argument("-o", "--output", help="Output PDF path (default: <input>_nup.pdf).")
    p.add_argument("-r", "--rows", type=int, required=True, help="Number of rows (m).")
    p.add_argument("-c", "--cols", type=int, required=True, help="Number of columns (n).")
    p.add_argument(
        "--orientation", choices=["portrait", "landscape"], default="portrait",
        help="Output page orientation.",
    )
    p.add_argument(
        "--page-size", default="A4",
        help="Named size (A4, A3, Letter, Legal, ...) or custom WxH (e.g. 210x297mm).",
    )
    p.add_argument(
        "--unit", choices=list(UNITS.keys()), default="mm",
        help="Unit for --margin, --gutter and bare custom sizes.",
    )
    p.add_argument("--margin", type=float, default=10.0, help="Outer margin.")
    p.add_argument("--gutter", type=float, default=5.0, help="Gap between cells.")
    p.add_argument(
        "--order", choices=["row", "col"], default="row",
        help="Fill order: row-major (left->right) or column-major (top->bottom).",
    )
    p.add_argument(
        "--rotate", choices=["0", "90", "180", "270", "auto"], default="0",
        help="Rotate each placed page by a fixed angle, or 'auto' to rotate "
             "90 when that makes the page fit larger (max possible size).",
    )
    p.add_argument(
        "--pages", default="all",
        help="Pages to include, 1-based, e.g. '1-10,15,20-'.",
    )
    p.add_argument("--frame", action="store_true", help="Draw a border around each placed page.")
    p.add_argument("--frame-width", type=float, default=0.5, help="Frame line width (pt).")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pdf_tool.py",
        description=(
            "PDF toolkit: N-up plus page operations (nup, merge, extract, delete, "
            "reorder, split, rotate, info). Page specs are 1-based, e.g. '1-3,5,8-'. "
            "Page operations write new files; inputs are never overwritten."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    sub = p.add_subparsers(dest="command", required=True)

    n = sub.add_parser(
        "nup",
        help="Combine multiple pages onto each sheet (N-up), auto-fit & centered.",
        description="Combine multiple PDF pages onto single pages (N-up), auto-fit & centered.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    add_nup_arguments(n)
    n.set_defaults(func=cmd_nup)

    m = sub.add_parser(
        "merge",
        help="Concatenate PDFs and/or images in order.",
        description=(
            "Concatenate PDFs and/or images in the given order. PDFs are appended "
            "as-is (optional per-file subset FILE:1-3,5). Each image becomes one page "
            "of the chosen size, auto-scaled to fit (aspect ratio kept) and centered."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    m.add_argument("inputs", nargs="+", help="Input PDFs/images in order. Per-file PDF subset: FILE:1-3,5")
    m.add_argument("-o", "--output", default="merged.pdf", help="Output PDF path.")
    img = m.add_argument_group("image inputs (apply to image inputs only)")
    img.add_argument(
        "--page-size", default="A4",
        help="Page size for images: named (A4, A3, Letter, ...) or custom WxH (e.g. 210x297mm).",
    )
    img.add_argument(
        "--orientation", choices=["portrait", "landscape", "auto"], default="auto",
        help="Page orientation for images; 'auto' matches each image's own orientation.",
    )
    img.add_argument(
        "--unit", choices=list(UNITS.keys()), default="mm",
        help="Unit for --margin and bare custom --page-size.",
    )
    img.add_argument("--margin", type=float, default=0.0, help="Margin around the image on its page.")
    m.set_defaults(func=cmd_merge)

    e = sub.add_parser("extract", help="Write a subset of pages to a new PDF.")
    e.add_argument("-i", "--input", required=True, help="Input PDF path.")
    e.add_argument("-o", "--output", help="Output PDF (default: <input>_extract.pdf).")
    e.add_argument("-p", "--pages", required=True, help="Pages to keep, e.g. '1-3,5,8-'.")
    e.set_defaults(func=cmd_extract)

    d = sub.add_parser("delete", help="Remove a subset of pages; keep the rest.")
    d.add_argument("-i", "--input", required=True, help="Input PDF path.")
    d.add_argument("-o", "--output", help="Output PDF (default: <input>_deleted.pdf).")
    d.add_argument("-p", "--pages", required=True, help="Pages to remove, e.g. '2,4,6-8'.")
    d.set_defaults(func=cmd_delete)

    r = sub.add_parser(
        "reorder",
        help="Rearrange pages: small edits (--swap/--move) or full order (--order/--reverse).",
        description=(
            "Rearrange pages. For a few edits in a large PDF, prefer --swap/--move "
            "(they start from the existing order, so you only specify what changes). "
            "Use --order for an explicit permutation (ranges allowed, e.g. "
            "'500,1-499,501-') or --reverse for the whole document."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    r.add_argument("-i", "--input", required=True, help="Input PDF path.")
    r.add_argument("-o", "--output", help="Output PDF (default: <input>_reordered.pdf).")
    r.add_argument(
        "--swap", action="append", metavar="A:B",
        help="Swap pages A and B (1-based, original numbers). Repeatable or comma-separated, "
             "e.g. '3:7,10:20'. Applied simultaneously; pairs must be disjoint.",
    )
    r.add_argument(
        "--move", action="append", metavar="SRC:DEST",
        help="Move page/block SRC so it starts at position DEST. SRC is a page ('50') or "
             "block ('5-8'); DEST is a 1-based position or 'start'/'end'. Repeatable/"
             "comma-separated; applied left-to-right on the current order.",
    )
    r.add_argument("--order", help="Explicit new order, e.g. '3,1,2,4-' (permutation unless --allow-partial).")
    r.add_argument("--reverse", action="store_true", help="Reverse the whole page order.")
    r.add_argument("--allow-partial", action="store_true", help="Allow dropping/duplicating pages in --order.")
    r.set_defaults(func=cmd_reorder)

    s = sub.add_parser("split", help="Split into multiple files (per N pages or by ranges).")
    s.add_argument("-i", "--input", required=True, help="Input PDF path.")
    s.add_argument("--outdir", help="Output directory (default: input's directory).")
    grp = s.add_mutually_exclusive_group()
    grp.add_argument("--every", type=int, default=1, help="Pages per output file.")
    grp.add_argument("--ranges", help="Explicit ranges, one file each, e.g. '1-3,4-6,7-'.")
    s.set_defaults(func=cmd_split)

    ro = sub.add_parser(
        "rotate",
        help="Rotate the whole file or a subset of pages by a multiple of 90 degrees.",
        description=(
            "Rotate by a multiple of 90 degrees - the whole file (default) or a "
            "subset via -p/--pages. Give the angle with --angle, or use a shortcut: "
            "--cw (90), --ccw (270), --flip (180). Rotation is relative to each "
            "page's current angle unless --absolute is given."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    ro.add_argument("-i", "--input", required=True, help="Input PDF path.")
    ro.add_argument("-o", "--output", help="Output PDF (default: <input>_rotated.pdf).")
    angle_grp = ro.add_mutually_exclusive_group(required=True)
    angle_grp.add_argument("--angle", type=int, help="Any multiple of 90 (e.g. 90, 180, 270, -90).")
    angle_grp.add_argument(
        "--cw", "--right", dest="shortcut", action="store_const", const=90,
        help="Shortcut for 90 degrees clockwise.",
    )
    angle_grp.add_argument(
        "--ccw", "--left", dest="shortcut", action="store_const", const=270,
        help="Shortcut for 90 degrees counter-clockwise (= 270).",
    )
    angle_grp.add_argument(
        "--flip", "--upside-down", dest="shortcut", action="store_const", const=180,
        help="Shortcut for 180 degrees (upside down).",
    )
    ro.add_argument("-p", "--pages", default="all", help="Pages to rotate (default: all = whole file).")
    ro.add_argument("--absolute", action="store_true", help="Set absolute rotation instead of adding to current.")
    ro.set_defaults(func=cmd_rotate, angle=None, shortcut=None)

    inf = sub.add_parser("info", help="Show page count, sizes, rotation and metadata.")
    inf.add_argument("-i", "--input", required=True, help="Input PDF path.")
    inf.add_argument("--per-page", action="store_true", help="List each page's size and rotation.")
    inf.set_defaults(func=cmd_info)

    bm = sub.add_parser(
        "bookmark",
        help="Read or edit PDF bookmarks (outline / table of contents).",
        description=(
            "Read or edit PDF bookmarks (the outline / table of contents). Actions: "
            "list, add, delete, update, export, import. Bookmarks form a tree via "
            "1-based levels: the first must be level 1 and each level may only step up "
            "by 1. Editing actions write a new file; the input is never overwritten."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    bsub = bm.add_subparsers(dest="bookmark_command", required=True)

    bl = bsub.add_parser(
        "list",
        help="Print the current bookmarks.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    bl.add_argument("-i", "--input", required=True, help="Input PDF path.")
    bl.add_argument("--json", action="store_true", help="Output as JSON instead of an indented tree.")
    bl.set_defaults(func=cmd_bm_list)

    ba = bsub.add_parser(
        "add",
        help="Add one or more bookmarks.",
        description=(
            "Add one or more bookmarks (repeat --add for several). By default each new "
            "bookmark is inserted so the list stays ordered by page; use --at to place "
            "them at an explicit position. For complex nested trees use 'bookmark import'."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    ba.add_argument("-i", "--input", required=True, help="Input PDF path.")
    ba.add_argument("-o", "--output", help="Output PDF (default: <input>_bookmarks.pdf).")
    ba.add_argument(
        "--add", action="append", nargs=3, metavar=("PAGE", "LEVEL", "TITLE"),
        help="Add a bookmark to 1-based PAGE at hierarchy LEVEL (1=top) titled TITLE. "
             "Repeatable; quote TITLE if it has spaces. Use PAGE -1 for no destination.",
    )
    ba.add_argument(
        "--at", type=int,
        help="Insert the new bookmark(s) at this 1-based position instead of by page order.",
    )
    ba.set_defaults(func=cmd_bm_add)

    bd = bsub.add_parser(
        "delete",
        help="Delete bookmarks by index, or all of them.",
        description=(
            "Delete bookmarks selected by --index (1-based, e.g. '2,4-6'), or all of "
            "them with --all. By default a deleted bookmark also removes its children "
            "(its whole subtree); use --keep-children to promote children up one level."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    bd.add_argument("-i", "--input", required=True, help="Input PDF path.")
    bd.add_argument("-o", "--output", help="Output PDF (default: <input>_bookmarks.pdf).")
    bd.add_argument("--index", help="1-based bookmark index/indices to delete, e.g. '3' or '2,4-6'.")
    bd.add_argument("--all", action="store_true", help="Delete every bookmark (clear the outline).")
    bd.add_argument(
        "--keep-children", action="store_true",
        help="Promote children instead of deleting the whole subtree.",
    )
    bd.set_defaults(func=cmd_bm_delete)

    bu = bsub.add_parser(
        "update",
        help="Change title/page/level of existing bookmarks.",
        description=(
            "Update existing bookmarks selected by --index (1-based, e.g. '2,4-6'); "
            "apply any of --title/--page/--level to all selected entries. For distinct "
            "per-entry edits, use 'bookmark export', edit the JSON, then 'bookmark import'."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    bu.add_argument("-i", "--input", required=True, help="Input PDF path.")
    bu.add_argument("-o", "--output", help="Output PDF (default: <input>_bookmarks.pdf).")
    bu.add_argument("--index", required=True, help="1-based bookmark index/indices to update, e.g. '3' or '2,4-6'.")
    bu.add_argument("--title", help="New title.")
    bu.add_argument("--page", type=int, help="New 1-based page (or -1 for no destination).")
    bu.add_argument("--level", type=int, help="New hierarchy level (1=top).")
    bu.set_defaults(func=cmd_bm_update)

    be = bsub.add_parser(
        "export",
        help="Export bookmarks to a JSON file.",
        description="Export the bookmarks as JSON (default: <input>_bookmarks.json; use -o - for stdout).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    be.add_argument("-i", "--input", required=True, help="Input PDF path.")
    be.add_argument("-o", "--output", help="Output JSON path, or '-' for stdout (default: <input>_bookmarks.json).")
    be.add_argument(
        "--details", action="store_true",
        help="Include destination details (position/zoom) for higher-fidelity round-trips.",
    )
    be.set_defaults(func=cmd_bm_export)

    bi = bsub.add_parser(
        "import",
        help="Replace bookmarks from a JSON file.",
        description=(
            "Replace ALL bookmarks with the contents of a JSON file (as produced by "
            "'bookmark export'). Entries may be objects {level,title,page[,dest]} or "
            "[level, title, page] arrays."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    bi.add_argument("-i", "--input", required=True, help="Input PDF path.")
    bi.add_argument("-o", "--output", help="Output PDF (default: <input>_bookmarks.pdf).")
    bi.add_argument("--from", dest="source", required=True, metavar="FILE", help="JSON file to import (or '-' for stdin).")
    bi.set_defaults(func=cmd_bm_import)

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    # Fill in default output for single-output commands, derived from the input.
    suffixes = {
        "nup": "_nup", "extract": "_extract", "delete": "_deleted",
        "reorder": "_reordered", "rotate": "_rotated",
    }
    if args.command in suffixes and not getattr(args, "output", None):
        base, _ = os.path.splitext(args.input)
        args.output = f"{base}{suffixes[args.command]}.pdf"

    try:
        args.func(args)
    except (ValueError, FileNotFoundError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
