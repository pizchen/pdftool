# pdf_tool.py — PDF toolkit (N-up + page operations)

**English** | [中文](README.zh-CN.md)

A single PyMuPDF-based CLI with subcommands for **N-up** layout and everyday PDF
page operations: `nup`, `merge`, `extract`, `delete`, `reorder`, `split`,
`rotate`, `info`, and `bookmark` (read/edit the outline).

The **`nup`** subcommand combines multiple pages of a PDF onto single output
pages in an **m rows × n columns** grid. Each source page is scaled as large as
possible inside its grid cell while keeping its aspect ratio (**auto-fit**) and is
centered. You choose **portrait** or **landscape** for the output page.

## Install

### Recommended: uv (portable, locked deps)

The project ships a `pyproject.toml` + `uv.lock`, so dependencies are pinned and
reproducible across machines (Linux/macOS/Windows wheels are all locked).

```bash
# On any machine with uv installed (https://docs.astral.sh/uv/):
uv sync                       # creates .venv from uv.lock
uv run python pdf_tool.py -h  # run the tool
```

To move the tool to another machine, copy these files (NOT `.venv`):
`pdf_tool.py`, `office2pdf.py`, `pyproject.toml`, `uv.lock`, `README.md`.
Then run `uv sync`.

### Alternative: plain pip

```bash
pip install -r requirements.txt
python3 pdf_tool.py -h
```

## N-up (`pdf_tool.py nup`)

```bash
python3 pdf_tool.py nup -i INPUT.pdf [-o OUTPUT.pdf] -r ROWS -c COLS [options]
```

### Key options

| Option | Default | Meaning |
| --- | --- | --- |
| `-r, --rows` | (required) | rows per sheet (m) |
| `-c, --cols` | (required) | columns per sheet (n) |
| `--orientation` | `portrait` | `portrait` or `landscape` output page |
| `--page-size` | `A4` | named (`A4`,`A3`,`A5`,`Letter`,`Legal`,…) or custom `WxH` (e.g. `210x297mm`) |
| `--unit` | `mm` | unit for `--margin`, `--gutter`, bare custom sizes (`pt`,`mm`,`cm`,`in`) |
| `--margin` | `10` | outer margin |
| `--gutter` | `5` | gap between cells |
| `--order` | `row` | `row` (left→right) or `col` (top→bottom) fill order |
| `--rotate` | `0` | rotate each placed page `0/90/180/270`, or `auto` to rotate 90° when that fills the cell more (max possible size) |
| `--pages` | `all` | subset, 1-based, e.g. `1-10,15,20-` |
| `--frame` | off | draw a box/border around each placed page |

If `-o` is omitted the output is `<input>_nup.pdf`.

## Examples

```bash
# 2x2 (4-up), A4 landscape
python3 pdf_tool.py nup -i in.pdf -r 2 -c 2 --orientation landscape

# 3 rows x 2 cols on A3 portrait, 8mm margin, 4mm gap, with page borders
python3 pdf_tool.py nup -i in.pdf -r 3 -c 2 --page-size A3 --margin 8 --gutter 4 --frame

# 1x2 booklet-style, only pages 1-20, US Letter landscape
python3 pdf_tool.py nup -i in.pdf -r 1 -c 2 --page-size Letter --orientation landscape --pages 1-20

# Boxed + auto max size: rotate each page 90 only if it ends up larger
python3 pdf_tool.py nup -i in.pdf -r 3 -c 2 --orientation portrait --frame --rotate auto
```

## Office documents → PDF (`office2pdf.py`)

`pdf_tool.py` only works on PDFs. To use it on Office files (PPTX, DOCX, ODP,
ODT, XLSX, …), first convert them to PDF with the companion tool
**`office2pdf.py`**, then run whatever `pdf_tool.py` command you want. It is a
**pure converter** — it does the Office→PDF step (via headless **LibreOffice**)
and nothing else, leaving the next step entirely up to you.

```bash
# Convert a slide deck -> deck.pdf (next to the input)
python3 office2pdf.py deck.pptx

# Custom output name
python3 office2pdf.py report.docx -o report.pdf

# Batch several files into one folder
python3 office2pdf.py a.pptx b.docx c.odp --outdir pdfs/

# Then do whatever you like with the PDF, e.g. a 4-up handout:
python3 pdf_tool.py nup -i deck.pdf -r 2 -c 2 --orientation landscape --frame
```

Options: `-o PATH` (output file, single input only) **or** `--outdir DIR`
(folder for one or more outputs; defaults to each input's own folder),
`--soffice PATH` (LibreOffice binary), `--timeout SECONDS` (conversion timeout,
default 180), `--force` (overwrite an existing output). If LibreOffice isn't
found, the tool prints OS-specific install instructions and exits.

Why convert at all? PPTX slides are page-like but can't be embedded into a PDF
cell natively, and DOCX is **reflowable** — it stores no pages at all (page
breaks are computed at render time). So a layout engine must materialize real
pages first; LibreOffice does exactly that, producing a normal PDF.

## Page operations

The remaining `pdf_tool.py` subcommands cover everyday PDF page wrangling. All
page specs are 1-based (`1-3,5,8-`), and every command writes a **new** file — an
input is never overwritten.

```bash
# Merge several PDFs in the given order
python3 pdf_tool.py merge a.pdf b.pdf c.pdf -o all.pdf

# Merge with a per-file page subset (cover page 1 of one + pages 2.. of another)
python3 pdf_tool.py merge cover.pdf:1 body.pdf:2- -o report.pdf

# Merge images (each auto-scaled to fit its page, centered) - mix with PDFs freely
python3 pdf_tool.py merge scan1.jpg scan2.png cover.pdf -o bundle.pdf

# Force a uniform page for images (A4 portrait, 10mm margin)
python3 pdf_tool.py merge page1.png page2.tif -o album.pdf --orientation portrait --margin 10

# Extract a subset of pages into a new PDF
python3 pdf_tool.py extract -i in.pdf -p 1-3,7 -o subset.pdf

# Delete pages (keep the rest)
python3 pdf_tool.py delete -i in.pdf -p 2,4 -o trimmed.pdf

# Reorder pages
python3 pdf_tool.py reorder -i in.pdf --swap 3:7 --swap 10:20   # switch a few pages
python3 pdf_tool.py reorder -i in.pdf --move 500:1             # move page 500 to the front
python3 pdf_tool.py reorder -i in.pdf --move 5-8:end          # move a block to the end
python3 pdf_tool.py reorder -i in.pdf --order 3,1,2,4-         # explicit permutation
python3 pdf_tool.py reorder -i in.pdf --reverse
# (use --allow-partial to intentionally drop/duplicate pages with --order)

# Split into multiple files
python3 pdf_tool.py split -i in.pdf --every 2            # chunks of 2 pages
python3 pdf_tool.py split -i in.pdf --ranges 1-3,4-6,7-  # one file per range

# Rotate: shortcuts --cw (90), --ccw (270), --flip (180), or --angle for any multiple of 90
python3 pdf_tool.py rotate -i in.pdf --cw                 # whole file, 90 clockwise
python3 pdf_tool.py rotate -i in.pdf --ccw -p 2,4-6       # subset, 90 counter-clockwise
python3 pdf_tool.py rotate -i in.pdf --flip               # 180 (upside down)
python3 pdf_tool.py rotate -i in.pdf --angle 270 -p 1     # explicit angle (relative by default)

# Inspect a PDF
python3 pdf_tool.py info -i in.pdf --per-page
```

| Command | Purpose | Key options |
| --- | --- | --- |
| `merge` | Concatenate PDFs and/or images in order | `inputs...` (PDF subset `FILE:pages`); images: `--page-size`/`--orientation`/`--unit`/`--margin`; `-o` |
| `extract` | Keep a subset of pages | `-i`, `-p/--pages`, `-o` |
| `delete` | Remove a subset of pages | `-i`, `-p/--pages`, `-o` |
| `reorder` | Rearrange pages | `-i`, `--swap` / `--move` / `--order` / `--reverse`, `--allow-partial`, `-o` |
| `split` | One PDF → many | `-i`, `--every N` \| `--ranges`, `--outdir` |
| `rotate` | Rotate whole file or subset | `-i`, `--cw`/`--ccw`/`--flip` or `--angle`, `-p/--pages`, `--absolute`, `-o` |
| `info` | Show pages/sizes/metadata | `-i`, `--per-page` |
| `bookmark` | Read/edit the outline (TOC) | `list` / `add` / `delete` / `update` / `export` / `import` (see [Bookmarks](#bookmarks-pdf_toolpy-bookmark)) |

### Reordering a few pages in a large PDF

You should **never** have to type out all pages. For small edits to a big file,
use the delta-based options — untouched pages keep their place:

- **Swap pages** (by original number, applied simultaneously, must be disjoint):

```bash
python3 pdf_tool.py reorder -i big.pdf --swap 3:7,10:20
```

- **Move a page or block** to a new spot (`DEST` is a 1-based position, or `start`/`end`):

```bash
python3 pdf_tool.py reorder -i big.pdf --move 500:1      # page 500 -> position 1
python3 pdf_tool.py reorder -i big.pdf --move 12-15:300  # block 12-15 starts at position 300
python3 pdf_tool.py reorder -i big.pdf --move 800:end
```

`--swap`/`--move` are repeatable and comma-separated. Swaps are computed together;
moves are applied left-to-right on the current order. If you need an explicit
permutation instead, `--order` accepts ranges and open ranges, so it also stays
compact — e.g. moving page 500 to the front of a 1000-page file is just
`--order 500,1-499,501-`.

### Merging images into a PDF

`merge` accepts image files alongside (or instead of) PDFs. Each image becomes
one page of `--page-size` (default `A4`), and is **auto-scaled to fit** the page
with its aspect ratio preserved and **centered**. `--orientation auto` (default)
gives each image a page matching its own orientation; use `portrait`/`landscape`
for uniform pages, and `--margin` for a border. A `FILE:pages` subset is only
valid for PDFs, not images.

- **Supported natively** (no extra dependency): JPEG, PNG, GIF, BMP, TIFF,
  JPEG2000, PNM/PGM/PBM/PPM — decoded directly by PyMuPDF (original compression
  kept).
- **WEBP / HEIC and similar**: handled only if [Pillow](https://python-pillow.org/)
  is installed (`pip install pillow`; add `pillow-heif` for HEIC). Without Pillow
  these report a clear error. Pillow is **optional** and not part of the locked
  dependencies.

Run `python3 pdf_tool.py <command> -h` for the full option list of any command.

## Bookmarks (`pdf_tool.py bookmark`)

Read and edit a PDF's **bookmarks** (the outline / table of contents) through
the `bookmark` subcommand. Bookmarks form a **tree** encoded by a 1-based
`level`: the first bookmark must be level 1 and each level may only step up by 1
(`1 -> 2`, never `1 -> 3`). Pages are 1-based; use `-1` for a bookmark with no
page destination. Every editing action writes a **new** file — the input is
never overwritten (default output `<input>_bookmarks.pdf`).

```bash
# Read the current bookmarks (indented tree with 1-based indices)
python3 pdf_tool.py bookmark list -i in.pdf
# ... or as JSON
python3 pdf_tool.py bookmark list -i in.pdf --json

# Add one or more bookmarks: --add PAGE LEVEL "TITLE" (repeatable)
python3 pdf_tool.py bookmark add -i in.pdf \
    --add 1 1 "Cover" --add 2 1 "Chapter 1" --add 3 2 "Section 1.1"
# Insert at an explicit position instead of by page order
python3 pdf_tool.py bookmark add -i in.pdf --add 10 1 "Appendix" --at 1

# Delete by 1-based index (a deleted bookmark takes its children with it)
python3 pdf_tool.py bookmark delete -i in.pdf --index 2,4-6
python3 pdf_tool.py bookmark delete -i in.pdf --index 2 --keep-children  # promote kids
python3 pdf_tool.py bookmark delete -i in.pdf --all                      # clear all

# Update title/page/level of selected bookmarks
python3 pdf_tool.py bookmark update -i in.pdf --index 3 --title "New title"
python3 pdf_tool.py bookmark update -i in.pdf --index 2,5 --level 2

# Export to JSON, edit it, then import to replace the whole outline
python3 pdf_tool.py bookmark export -i in.pdf -o toc.json
python3 pdf_tool.py bookmark import -i in.pdf --from toc.json -o out.pdf
```

| Action | Purpose | Key options |
| --- | --- | --- |
| `list` | Print the outline (indented tree or JSON) | `-i`, `--json` |
| `add` | Add one or more bookmarks | `-i`, `--add PAGE LEVEL TITLE` (repeatable), `--at`, `-o` |
| `delete` | Remove bookmarks (subtree by default) | `-i`, `--index` \| `--all`, `--keep-children`, `-o` |
| `update` | Change title/page/level | `-i`, `--index`, `--title`/`--page`/`--level`, `-o` |
| `export` | Dump the outline to JSON | `-i`, `-o` (`-` for stdout), `--details` |
| `import` | Replace the outline from JSON | `-i`, `--from FILE` (`-` for stdin), `-o` |

For quick single or few edits, use `add`/`delete`/`update` directly. For building
or restructuring a **deeply nested** outline, prefer the `export` -> edit ->
`import` round-trip: export writes one JSON object `{level, title, page}` per
bookmark (add `--details` to also capture each destination's position/zoom), and
`import` validates and replaces the entire outline in one shot.

## Dependencies

| Component | What it needs | Managed by |
| --- | --- | --- |
| `pdf_tool.py` (all subcommands) | `pymupdf` only | `requirements.txt` / `pyproject.toml` / `uv.lock` |
| `office2pdf.py` | **Python standard library only** (no pip package at all) | — |
| `office2pdf.py`, conversion | **LibreOffice** (`soffice` binary) | **system** install — *not* pip/uv |

The only Python dependency is `pymupdf` (for `pdf_tool.py`); `office2pdf.py` needs
no pip package whatsoever. LibreOffice is a **system** dependency: `pip`/`uv`
cannot lock a native binary, so it is intentionally kept out of `requirements.txt`
and `uv.lock`. `office2pdf.py` checks for it at runtime and prints OS-specific
install instructions if it's missing.

```bash
# Install LibreOffice (only for Office inputs)
sudo apt install libreoffice          # Debian/Ubuntu
brew install --cask libreoffice       # macOS
```

## Notes

- **Box per page** (`--frame`): the border is drawn tightly around each placed
  page, so it doubles as a visual outline of the original page boundary.
- **Max possible size** is always applied within each cell (aspect ratio kept,
  centered). Adding `--rotate auto` additionally rotates a page 90° whenever the
  rotated orientation fits larger — useful when source and output orientations
  differ (e.g. wide slides on a portrait sheet).
