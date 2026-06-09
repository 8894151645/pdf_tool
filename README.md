# 📄 PDF 工具箱

一个用 Python 写的 PDF 工具箱，提供 **命令行版** 和 **图形界面版**。适合办公场景，用来处理 PDF 合并、拆分、页面调整和常见格式转换。

## 功能列表

| 模块 | 功能 |
|------|------|
| 页面组织 | 合并 PDF、拆分 PDF、删除页面、提取页面、调整页序、旋转页面 |
| 格式转换 | PDF → 图片、图片 → PDF、PDF → Word、Word/PPT → PDF、PDF → Excel、PDF → 纯文本 |

## 文件说明

| 文件 | 说明 |
|------|------|
| `pdf_toolbox.py` | 命令行主程序 |
| `pdf_toolbox_gui.py` | 图形界面程序 |
| `requirements.txt` | Python 依赖清单 |
| `install.bat` | Windows 双击自动安装依赖，脚本内容使用英文，避免中文编码导致的问题 |
| `run_gui.bat` | Windows 双击启动图形界面，脚本内容使用英文，避免中文编码导致的问题 |

> 你提到的 `dat` 文件，如果目的是“双击自动安装”，Windows 上通常应使用 `.bat` 批处理文件。因此本项目提供的是 `install.bat`。

## Windows 快速使用

### 1. 安装 Python

推荐安装 Python 3.9 或更高版本；**Python 3.8 通常也可以使用**。

下载地址：

https://www.python.org/downloads/

安装时建议勾选：

```text
Add Python to PATH
```

### 2. 双击安装依赖

下载项目后，双击：

```text
install.bat
```

它会自动执行：

```bash
python -m pip install -r requirements.txt
```

`install.bat` 内部提示文字已改为英文，减少 Windows 批处理中文编码乱码或执行异常的概率。

### 3. 双击启动图形界面

依赖安装完成后，双击：

```text
run_gui.bat
```

或者手动执行：

```bash
python pdf_toolbox_gui.py
```

`run_gui.bat` 内部提示文字同样使用英文。

## 图形界面使用说明

打开界面后可以在两个标签页中选择功能：

### 页面组织

- **合并 PDF**：在“多文件输入”里添加多个 PDF，选择输出 PDF 路径后执行。
- **拆分 PDF**：选择输入 PDF、输出目录，并填写“拆分页数”。
- **删除页面**：选择输入 PDF、输出 PDF，并填写页码，例如 `2,4,6-8`。
- **提取页面**：选择输入 PDF、输出 PDF，并填写页码，例如 `1,3,5-7`。
- **调整页序**：选择输入 PDF、输出 PDF，并填写新顺序，例如 `3,1,2`。
- **旋转页面**：选择输入 PDF、输出 PDF，选择旋转角度；页码为空表示旋转全部页面。

### 格式转换

- **PDF → 图片**：选择输入 PDF、输出目录，可设置 DPI 和图片格式。
- **图片 → PDF**：在“多图片输入”里添加图片，选择输出 PDF 路径后执行。
- **PDF → Word**：选择输入 PDF 和输出 `.docx` 文件路径。
- **Word/PPT → PDF**：选择 Word/PPT 文件和输出目录。
- **PDF → Excel**：选择输入 PDF 和输出 `.xlsx` 文件路径。
- **PDF → 纯文本**：选择输入 PDF 和输出 `.txt` 文件路径。

## 命令行用法速查

| 功能 | 命令示例 |
|------|---------|
| 合并 | `python pdf_toolbox.py merge a.pdf b.pdf -o out.pdf` |
| 拆分 | `python pdf_toolbox.py split in.pdf -s 2 -o 输出目录` |
| 删除页面 | `python pdf_toolbox.py delete in.pdf -p 2,4 -o out.pdf` |
| 提取页面 | `python pdf_toolbox.py extract in.pdf -p 1,3-5 -o out.pdf` |
| 调整页序 | `python pdf_toolbox.py reorder in.pdf -r 3,1,2 -o out.pdf` |
| 旋转 | `python pdf_toolbox.py rotate in.pdf -a 90 -o out.pdf` |
| PDF→图片 | `python pdf_toolbox.py pdf2img in.pdf --dpi 200` |
| 图片→PDF | `python pdf_toolbox.py img2pdf 1.jpg 2.jpg -o out.pdf` |
| PDF→Word | `python pdf_toolbox.py pdf2word in.pdf -o out.docx` |
| Office→PDF | `python pdf_toolbox.py office2pdf in.docx` |
| PDF→Excel | `python pdf_toolbox.py pdf2excel in.pdf -o out.xlsx` |
| PDF→文本 | `python pdf_toolbox.py pdf2text in.pdf -o out.txt` |

> 页码从 1 开始，支持 `2,4,6-8` 这种写法。查看帮助：`python pdf_toolbox.py -h`

## 重要说明

- **Python 版本**：推荐 Python 3.9+；如果你已经安装 Python 3.8，可以先直接使用，依赖安装失败时再考虑升级。
- **PDF→Word / PDF→Excel** 对扫描件或图片型 PDF 效果有限，这类文件需要 OCR。
- **Word/PPT→PDF** 需要额外安装 [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)。
- `install.bat` 只安装 Python 依赖，不会自动安装 LibreOffice。

## 技术栈

- [tkinter](https://docs.python.org/zh-cn/3/library/tkinter.html) — 图形界面，Python 自带
- [pypdf](https://github.com/py-pdf/pypdf) — 页面组织
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) — PDF 转图片/文本
- [Pillow](https://github.com/python-pillow/Pillow) — 图片转 PDF
- [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) — PDF 转 Word
- [pdfplumber](https://github.com/jsvine/pdfplumber) + [openpyxl](https://openpyxl.readthedocs.io/) — PDF 转 Excel
