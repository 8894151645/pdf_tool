# 📄 PDF 工具箱（命令行版）

一个用 Python 写的命令行 PDF 工具箱，集成 12 个常用功能，办公场景一站式解决。

## 功能列表

| 模块 | 功能 |
|------|------|
| 页面组织 | 合并、拆分、删除页面、提取页面、调整页序、旋转 |
| 格式转换 | PDF↔图片、PDF→Word、Office→PDF、PDF→Excel、PDF→文本 |

## 安装

```bash
pip install -r requirements.txt
```

> `office2pdf` 功能需额外安装 [LibreOffice](https://www.libreoffice.org/)。

## 用法速查

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

## 说明

- **页码从 1 开始**，支持 `2,4,6-8` 这种范围写法。
- **PDF→Word / PDF→Excel** 对扫描件（图片型 PDF）无效，需要先做 OCR。
- **Office→PDF** 依赖 LibreOffice，跨平台且免费。

## 技术栈

- [pypdf](https://github.com/py-pdf/pypdf) — 页面组织
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) — PDF 转图片/文本
- [Pillow](https://github.com/python-pillow/Pillow) — 图片转 PDF
- [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) — PDF 转 Word
- [pdfplumber](https://github.com/jsvine/pdfplumber) + [openpyxl](https://openpyxl.readthedocs.io/) — PDF 转 Excel
