# pdf_tool.py — PDF 工具箱（N-up 拼版 + 页面操作）

[English](README.md) | **中文**

一个基于 PyMuPDF 的命令行工具，包含用于 **N-up（多页合一）** 拼版和日常 PDF
页面操作的子命令：`nup`、`merge`、`extract`、`delete`、`reorder`、`split`、
`rotate`、`info`。

**`nup`** 子命令将 PDF 的多个页面按 **m 行 × n 列** 的网格合并到单个输出页面上。
每个源页面在保持其纵横比的前提下，被尽可能放大以填满所在的网格单元（**自动适配**），
并居中显示。你可以选择输出页面为 **纵向（portrait）** 或 **横向（landscape）**。

## 安装

### 推荐：uv（可移植，依赖锁定）

项目附带 `pyproject.toml` + `uv.lock`，因此依赖被固定且可在不同机器上复现
（Linux/macOS/Windows 的 wheel 均已锁定）。

```bash
# 在任何已安装 uv 的机器上（https://docs.astral.sh/uv/）：
uv sync                       # 根据 uv.lock 创建 .venv
uv run python pdf_tool.py -h  # 运行工具
```

要把工具迁移到另一台机器，复制以下文件（**不要**复制 `.venv`）：
`pdf_tool.py`、`office2pdf.py`、`pyproject.toml`、`uv.lock`、`README.md`。
然后运行 `uv sync`。

### 备选：使用 pip

```bash
pip install -r requirements.txt
python3 pdf_tool.py -h
```

## N-up 拼版（`pdf_tool.py nup`）

```bash
python3 pdf_tool.py nup -i INPUT.pdf [-o OUTPUT.pdf] -r ROWS -c COLS [options]
```

### 主要选项

| 选项 | 默认值 | 含义 |
| --- | --- | --- |
| `-r, --rows` | （必填） | 每张纸的行数（m） |
| `-c, --cols` | （必填） | 每张纸的列数（n） |
| `--orientation` | `portrait` | 输出页面为 `portrait`（纵向）或 `landscape`（横向） |
| `--page-size` | `A4` | 预设尺寸（`A4`、`A3`、`A5`、`Letter`、`Legal`…）或自定义 `WxH`（如 `210x297mm`） |
| `--unit` | `mm` | `--margin`、`--gutter` 及裸自定义尺寸所用的单位（`pt`、`mm`、`cm`、`in`） |
| `--margin` | `10` | 外边距 |
| `--gutter` | `5` | 单元格之间的间隙 |
| `--order` | `row` | 填充顺序：`row`（从左到右）或 `col`（从上到下） |
| `--rotate` | `0` | 将每个放置的页面旋转 `0/90/180/270`，或用 `auto` 在能填得更满时旋转 90°（取最大可能尺寸） |
| `--pages` | `all` | 页面子集，从 1 开始计数，例如 `1-10,15,20-` |
| `--frame` | 关闭 | 在每个放置的页面周围绘制方框/边框 |

若省略 `-o`，输出文件为 `<input>_nup.pdf`。

## 示例

```bash
# 2x2（4 合 1），A4 横向
python3 pdf_tool.py nup -i in.pdf -r 2 -c 2 --orientation landscape

# 3 行 x 2 列，A3 纵向，8mm 边距，4mm 间隙，带页面边框
python3 pdf_tool.py nup -i in.pdf -r 3 -c 2 --page-size A3 --margin 8 --gutter 4 --frame

# 1x2 小册子风格，仅第 1-20 页，US Letter 横向
python3 pdf_tool.py nup -i in.pdf -r 1 -c 2 --page-size Letter --orientation landscape --pages 1-20

# 带边框 + 自动最大尺寸：仅当旋转后更大时才将每页旋转 90°
python3 pdf_tool.py nup -i in.pdf -r 3 -c 2 --orientation portrait --frame --rotate auto
```

## Office 文档 → PDF（`office2pdf.py`）

`pdf_tool.py` 只能处理 PDF。若要对 Office 文件（PPTX、DOCX、ODP、ODT、XLSX…）使用它，
请先用配套工具 **`office2pdf.py`** 将其转换为 PDF，然后再运行你需要的任何
`pdf_tool.py` 命令。该工具是一个**纯转换器**——它只完成 Office→PDF 这一步
（通过无界面的 **LibreOffice**），别的什么都不做，后续步骤完全由你决定。

```bash
# 转换一个幻灯片 -> deck.pdf（与输入文件同目录）
python3 office2pdf.py deck.pptx

# 自定义输出名称
python3 office2pdf.py report.docx -o report.pdf

# 批量转换多个文件到一个文件夹
python3 office2pdf.py a.pptx b.docx c.odp --outdir pdfs/

# 然后对 PDF 做任何你想做的操作，例如 4 合 1 讲义：
python3 pdf_tool.py nup -i deck.pdf -r 2 -c 2 --orientation landscape --frame
```

选项：`-o PATH`（输出文件，仅限单个输入）**或** `--outdir DIR`（一个或多个输出
文件的目录；默认与各输入文件同目录）、`--soffice PATH`（LibreOffice 可执行文件）、
`--timeout SECONDS`（转换超时，默认 180）、`--force`（覆盖已存在的输出）。若找不到
LibreOffice，工具会打印针对各操作系统的安装说明并退出。

为什么需要转换这一步？PPTX 幻灯片虽类似页面，却无法原生地嵌入到 PDF 单元格中；
而 DOCX 是**可重排（reflowable）**的——它根本不存储页面（分页是在渲染时才计算的）。
因此必须先由排版引擎生成真实的页面；LibreOffice 正是做这件事，生成一个普通的 PDF。

## 页面操作

`pdf_tool.py` 的其余子命令涵盖日常的 PDF 页面处理。所有页面区间都从 1 开始计数
（`1-3,5,8-`），且每个命令都会写入一个**新**文件——输入文件永远不会被覆盖。

```bash
# 按给定顺序合并多个 PDF
python3 pdf_tool.py merge a.pdf b.pdf c.pdf -o all.pdf

# 合并时为每个文件指定页面子集（某文件的封面第 1 页 + 另一文件的第 2 页起）
python3 pdf_tool.py merge cover.pdf:1 body.pdf:2- -o report.pdf

# 合并图片（每张自动缩放以适配其页面，并居中）——可与 PDF 自由混合
python3 pdf_tool.py merge scan1.jpg scan2.png cover.pdf -o bundle.pdf

# 为图片强制统一页面（A4 纵向，10mm 边距）
python3 pdf_tool.py merge page1.png page2.tif -o album.pdf --orientation portrait --margin 10

# 提取页面子集到一个新 PDF
python3 pdf_tool.py extract -i in.pdf -p 1-3,7 -o subset.pdf

# 删除页面（保留其余）
python3 pdf_tool.py delete -i in.pdf -p 2,4 -o trimmed.pdf

# 重新排序页面
python3 pdf_tool.py reorder -i in.pdf --swap 3:7 --swap 10:20   # 互换几页
python3 pdf_tool.py reorder -i in.pdf --move 500:1             # 将第 500 页移到最前
python3 pdf_tool.py reorder -i in.pdf --move 5-8:end          # 将一段移到最后
python3 pdf_tool.py reorder -i in.pdf --order 3,1,2,4-         # 显式排列
python3 pdf_tool.py reorder -i in.pdf --reverse
# （配合 --allow-partial 可在使用 --order 时有意丢弃/重复页面）

# 拆分为多个文件
python3 pdf_tool.py split -i in.pdf --every 2            # 每 2 页一块
python3 pdf_tool.py split -i in.pdf --ranges 1-3,4-6,7-  # 每个区间一个文件

# 旋转：快捷方式 --cw（90）、--ccw（270）、--flip（180），或用 --angle 指定任意 90 的倍数
python3 pdf_tool.py rotate -i in.pdf --cw                 # 整个文件，顺时针 90°
python3 pdf_tool.py rotate -i in.pdf --ccw -p 2,4-6       # 子集，逆时针 90°
python3 pdf_tool.py rotate -i in.pdf --flip               # 180°（上下颠倒）
python3 pdf_tool.py rotate -i in.pdf --angle 270 -p 1     # 显式角度（默认为相对旋转）

# 查看 PDF 信息
python3 pdf_tool.py info -i in.pdf --per-page
```

| 命令 | 用途 | 主要选项 |
| --- | --- | --- |
| `merge` | 按顺序拼接 PDF 和/或图片 | `inputs...`（PDF 子集 `FILE:pages`）；图片：`--page-size`/`--orientation`/`--unit`/`--margin`；`-o` |
| `extract` | 保留页面子集 | `-i`、`-p/--pages`、`-o` |
| `delete` | 删除页面子集 | `-i`、`-p/--pages`、`-o` |
| `reorder` | 重新排列页面 | `-i`、`--swap` / `--move` / `--order` / `--reverse`、`--allow-partial`、`-o` |
| `split` | 一个 PDF → 多个 | `-i`、`--every N` \| `--ranges`、`--outdir` |
| `rotate` | 旋转整个文件或子集 | `-i`、`--cw`/`--ccw`/`--flip` 或 `--angle`、`-p/--pages`、`--absolute`、`-o` |
| `info` | 显示页数/尺寸/元数据 | `-i`、`--per-page` |

### 在大型 PDF 中重排少量页面

你**绝不**需要把所有页面都列出来。对大文件做少量改动时，使用基于增量（delta）的
选项——未触及的页面保持原位：

- **互换页面**（按原始页码，同时生效，且各页码必须互不相交）：

```bash
python3 pdf_tool.py reorder -i big.pdf --swap 3:7,10:20
```

- **移动一页或一段**到新位置（`DEST` 是从 1 开始的位置，或 `start`/`end`）：

```bash
python3 pdf_tool.py reorder -i big.pdf --move 500:1      # 第 500 页 -> 位置 1
python3 pdf_tool.py reorder -i big.pdf --move 12-15:300  # 第 12-15 段从位置 300 开始
python3 pdf_tool.py reorder -i big.pdf --move 800:end
```

`--swap`/`--move` 可重复使用，并以逗号分隔。互换会被一起计算；移动则按当前顺序
从左到右依次应用。若你需要显式排列，`--order` 支持区间和开区间，因此同样紧凑——
例如把第 500 页移到 1000 页文件的最前面，只需 `--order 500,1-499,501-`。

### 将图片合并为 PDF

`merge` 接受图片文件与 PDF 并存（或仅图片）。每张图片成为一页 `--page-size`
（默认 `A4`），并**自动缩放适配**页面，保持其纵横比并**居中**。`--orientation auto`
（默认）会让每张图片的页面方向与其自身一致；使用 `portrait`/`landscape` 可获得统一
方向的页面，`--margin` 可加边距。`FILE:pages` 子集仅对 PDF 有效，对图片无效。

- **原生支持**（无需额外依赖）：JPEG、PNG、GIF、BMP、TIFF、JPEG2000、
  PNM/PGM/PBM/PPM——由 PyMuPDF 直接解码（保留原始压缩）。
- **WEBP / HEIC 等**：仅在安装了 [Pillow](https://python-pillow.org/) 时才支持
  （`pip install pillow`；HEIC 还需 `pillow-heif`）。若未安装 Pillow，会给出明确的
  错误提示。Pillow 是**可选的**，不属于锁定的依赖。

运行 `python3 pdf_tool.py <command> -h` 可查看任意命令的完整选项列表。

## 依赖

| 组件 | 所需 | 由谁管理 |
| --- | --- | --- |
| `pdf_tool.py`（所有子命令） | 仅需 `pymupdf` | `requirements.txt` / `pyproject.toml` / `uv.lock` |
| `office2pdf.py` | **仅 Python 标准库**（完全不需要任何 pip 包） | — |
| `office2pdf.py`，转换功能 | **LibreOffice**（`soffice` 可执行文件） | **系统**安装——*非* pip/uv |

唯一的 Python 依赖是 `pymupdf`（供 `pdf_tool.py` 使用）；`office2pdf.py` 完全不需要
任何 pip 包。LibreOffice 是**系统**依赖：`pip`/`uv` 无法锁定原生二进制文件，因此它被
有意排除在 `requirements.txt` 和 `uv.lock` 之外。`office2pdf.py` 会在运行时检查它，
若缺失则打印针对各操作系统的安装说明。

```bash
# 安装 LibreOffice（仅 Office 输入需要）
sudo apt install libreoffice          # Debian/Ubuntu
brew install --cask libreoffice       # macOS
```

## 说明

- **每页方框**（`--frame`）：边框紧贴每个放置的页面绘制，因此它同时也勾勒出原始
  页面的边界。
- 每个单元格内始终采用**最大可能尺寸**（保持纵横比，居中）。额外加上 `--rotate auto`
  时，只要旋转后的方向能放得更大，就会把页面旋转 90°——这在源页面与输出页面方向
  不同时很有用（例如把宽幻灯片放到纵向纸上）。
