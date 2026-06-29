#!/usr/bin/env python3
"""
office2pdf.py - Convert Office documents to PDF using headless LibreOffice.

A focused converter: it turns PPTX/DOCX/ODP/ODT/XLSX/... into PDF and nothing
else. It locates the LibreOffice ``soffice`` binary and, if it is missing, tells
you exactly how to install it. What you do with the resulting PDF afterwards
(e.g. N-up or page operations via ``pdf_tool.py``) is entirely up to you.

Dependencies:
  * Python: standard library only - no pip packages.
  * System: LibreOffice (``soffice``) on PATH (or passed via --soffice).

Examples:
    python3 office2pdf.py deck.pptx                      # -> deck.pdf (next to input)
    python3 office2pdf.py report.docx -o out.pdf         # custom output name
    python3 office2pdf.py a.pptx b.docx --outdir pdfs     # batch into a folder

Next step is yours, e.g.:
    python3 pdf_tool.py nup -i deck.pdf -r 2 -c 2 --orientation landscape
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

# Office/OpenDocument formats LibreOffice can render to PDF.
OFFICE_EXTS = {
    ".pptx", ".ppt", ".pps", ".ppsx", ".odp",  # presentations
    ".docx", ".doc", ".odt", ".rtf",           # text documents
    ".xlsx", ".xls", ".ods",                   # spreadsheets (best-effort layout)
}


def find_soffice(explicit: str | None) -> str:
    """Locate the LibreOffice binary, or raise with install instructions."""
    candidates = [explicit] if explicit else ["soffice", "libreoffice"]
    for cand in candidates:
        if cand:
            found = shutil.which(cand)
            if found:
                return found
    raise RuntimeError(
        "LibreOffice was not found on your PATH, but it is required to convert "
        "Office documents to PDF.\n"
        "Install it, then re-run:\n"
        "  - Debian/Ubuntu : sudo apt install libreoffice\n"
        "  - Fedora        : sudo dnf install libreoffice\n"
        "  - macOS         : brew install --cask libreoffice\n"
        "  - Windows       : winget install TheDocumentFoundation.LibreOffice\n"
        "Or point to an existing binary with --soffice /path/to/soffice."
    )


def convert_to_pdf(
    src_path: str,
    soffice: str,
    timeout: float,
    profile_dir: str,
    out_dir: str,
) -> str:
    """Convert one Office file to PDF via headless LibreOffice; return the path.

    A per-run user profile keeps the conversion hermetic and avoids clashing with
    any LibreOffice instance the user already has open.
    """
    os.makedirs(out_dir, exist_ok=True)
    cmd = [
        soffice, "--headless", "--norestore", "--nolockcheck",
        f"-env:UserInstallation=file://{profile_dir}",
        "--convert-to", "pdf", "--outdir", out_dir, src_path,
    ]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"LibreOffice conversion timed out after {timeout:.0f}s for {src_path!r}. "
            "Increase --timeout or check the file."
        )

    base = os.path.splitext(os.path.basename(src_path))[0]
    pdf_path = os.path.join(out_dir, base + ".pdf")
    if proc.returncode != 0 or not os.path.exists(pdf_path):
        detail = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(
            f"LibreOffice failed to convert {src_path!r} to PDF"
            + (f":\n{detail}" if detail else ".")
        )
    return pdf_path


def resolve_output(src: str, output: str | None, outdir: str | None) -> str:
    """Decide the final PDF path for a given input."""
    if output:
        return output
    target_dir = outdir or os.path.dirname(os.path.abspath(src))
    return os.path.join(target_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="office2pdf.py",
        description=(
            "Convert Office documents (PPTX/DOCX/ODP/ODT/XLSX/...) to PDF using "
            "headless LibreOffice. Conversion only - use pdf_tool.py afterwards for "
            "any PDF operations (N-up, merge, extract, rotate, ...)."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("inputs", nargs="+", help="Office file(s) to convert to PDF.")
    dest = p.add_mutually_exclusive_group()
    dest.add_argument("-o", "--output", help="Output PDF path (only valid with a single input).")
    dest.add_argument("--outdir", help="Directory for the PDF(s) (default: each input's own folder).")
    p.add_argument("--soffice", help="Path to the LibreOffice 'soffice' binary.")
    p.add_argument("--timeout", type=float, default=180.0, help="Max seconds per conversion.")
    p.add_argument("--force", action="store_true", help="Overwrite an existing output PDF.")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.output and len(args.inputs) != 1:
        print(
            "Error: -o/--output works with a single input; use --outdir for "
            "multiple files.",
            file=sys.stderr,
        )
        return 1

    try:
        soffice = find_soffice(args.soffice)
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    tmp_root = tempfile.mkdtemp(prefix="office2pdf_")
    profile_dir = os.path.join(tmp_root, "profile")
    converted: list[str] = []
    try:
        for idx, src in enumerate(args.inputs):
            if not os.path.exists(src):
                raise FileNotFoundError(f"Input file not found: {src}")
            ext = os.path.splitext(src)[1].lower()
            if ext not in OFFICE_EXTS:
                supported = ", ".join(sorted(OFFICE_EXTS))
                raise ValueError(
                    f"Unsupported input '{src}' (extension '{ext}'). This tool "
                    f"converts Office files to PDF. Supported: {supported}."
                )

            final = resolve_output(src, args.output, args.outdir)
            if os.path.abspath(final) == os.path.abspath(src):
                raise ValueError(f"Refusing to overwrite the input file: {src}")
            if os.path.exists(final) and not args.force:
                raise ValueError(f"Output '{final}' already exists; use --force to overwrite.")

            out_dir = os.path.join(tmp_root, f"out{idx}")
            tmp_pdf = convert_to_pdf(src, soffice, args.timeout, profile_dir, out_dir)
            os.makedirs(os.path.dirname(os.path.abspath(final)), exist_ok=True)
            shutil.move(tmp_pdf, final)
            converted.append(final)
            print(f"Converted: {src} -> {final}")
    except (ValueError, FileNotFoundError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)

    print(f"Done: {len(converted)} file(s) converted to PDF.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
