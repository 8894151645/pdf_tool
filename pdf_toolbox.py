#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 工具箱（命令行版）
功能：合并 / 拆分 / 删除页面 / 提取页面 / 调整页序 / 旋转 /
      PDF↔图片 / PDF→Word / Office→PDF / PDF→Excel / PDF→文本
用法：python pdf_toolbox.py <命令> [参数]，详见 python pdf_toolbox.py -h
"""

import argparse
import os
import sys
from pathlib import Path


# ==================== 辅助函数 ====================

def parse_pages(pages_str):
    """把 '1,3,5-7' 解析成 [1, 3, 5, 6, 7]"""
    result = []
    for part in pages_str.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            start, end = part.split('-', 1)
            result.extend(range(int(start), int(end) + 1))
        else:
            result.append(int(part))
    return result


def ensure_dir(path):
    """确保输出目录存在"""
    Path(path).mkdir(parents=True, exist_ok=True)


# ==================== 1. 页面组织（pypdf）====================

def cmd_merge(args):
    from pypdf import PdfWriter
    writer = PdfWriter()
    for path in args.inputs:
        print(f"  添加: {path}")
        writer.append(path)
    writer.write(args.output)
    writer.close()
    print(f"✅ 合并完成 -> {args.output}")


def cmd_split(args):
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(args.input)
    total = len(reader.pages)
    ensure_dir(args.output_dir)
    stem = Path(args.input).stem
    count = 0
    for start in range(0, total, args.size):
        writer = PdfWriter()
        for i in range(start, min(start + args.size, total)):
            writer.add_page(reader.pages[i])
        count += 1
        out = os.path.join(args.output_dir, f"{stem}_part{count}.pdf")
        writer.write(out)
        print(f"  生成: {out}")
    print(f"✅ 拆分完成，共 {count} 个文件 -> {args.output_dir}")


def cmd_delete(args):
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(args.input)
    to_delete = set(parse_pages(args.pages))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages, start=1):
        if i not in to_delete:
            writer.add_page(page)
    writer.write(args.output)
    print(f"✅ 已删除页面 {sorted(to_delete)} -> {args.output}")


def cmd_extract(args):
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(args.input)
    pages = parse_pages(args.pages)
    writer = PdfWriter()
    for p in pages:
        writer.add_page(reader.pages[p - 1])
    writer.write(args.output)
    print(f"✅ 已提取页面 {pages} -> {args.output}")


def cmd_reorder(args):
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(args.input)
    order = parse_pages(args.order)
    writer = PdfWriter()
    for p in order:
        writer.add_page(reader.pages[p - 1])
    writer.write(args.output)
    print(f"✅ 已按新顺序排列 -> {args.output}")


def cmd_rotate(args):
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(args.input)
    target = set(parse_pages(args.pages)) if args.pages else None
    writer = PdfWriter()
    for i, page in enumerate(reader.pages, start=1):
        if target is None or i in target:
            page.rotate(args.angle)
        writer.add_page(page)
    writer.write(args.output)
    scope = f"第 {sorted(target)} 页" if target else "全部页面"
    print(f"✅ 已旋转 {scope} {args.angle}° -> {args.output}")


# ==================== 2. 格式转换 ====================

def cmd_pdf2img(args):
    import fitz  # PyMuPDF
    doc = fitz.open(args.input)
    ensure_dir(args.output_dir)
    stem = Path(args.input).stem
    n = len(doc)
    for i in range(n):
        pix = doc[i].get_pixmap(dpi=args.dpi)
        out = os.path.join(args.output_dir, f"{stem}_p{i + 1}.{args.format}")
        pix.save(out)
        print(f"  生成: {out}")
    doc.close()
    print(f"✅ PDF 转图片完成，共 {n} 页 -> {args.output_dir}")


def cmd_img2pdf(args):
    from PIL import Image
    images = []
    for p in args.inputs:
        img = Image.open(p)
        images.append(img.convert('RGB') if img.mode != 'RGB' else img)
    images[0].save(args.output, save_all=True, append_images=images[1:])
    print(f"✅ {len(images)} 张图片合成 PDF -> {args.output}")


def cmd_pdf2word(args):
    from pdf2docx import Converter
    cv = Converter(args.input)
    cv.convert(args.output)
    cv.close()
    print(f"✅ PDF 转 Word 完成 -> {args.output}")


def cmd_office2pdf(args):
    import subprocess, shutil
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        print("❌ 未找到 LibreOffice，请先安装：https://www.libreoffice.org/")
        return
    ensure_dir(args.output_dir)
    subprocess.run([soffice, '--headless', '--convert-to', 'pdf',
                    '--outdir', args.output_dir, args.input], check=True)
    print(f"✅ Office 转 PDF 完成 -> {args.output_dir}")


def cmd_pdf2excel(args):
    import pdfplumber
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    found = 0
    with pdfplumber.open(args.input) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            for tno, table in enumerate(page.extract_tables(), start=1):
                found += 1
                ws = wb.create_sheet(f"P{pno}_T{tno}")
                for row in table:
                    ws.append([("" if c is None else c) for c in row])
    if found == 0:
        print("⚠️ 未在 PDF 中检测到表格")
        return
    wb.save(args.output)
    print(f"✅ 共提取 {found} 个表格 -> {args.output}")


def cmd_pdf2text(args):
    import fitz
    doc = fitz.open(args.input)
    text = [page.get_text() for page in doc]
    doc.close()
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write("\n".join(text))
    print(f"✅ 提取文字完成 -> {args.output}")


# ==================== 命令行定义 ====================

def build_parser():
    parser = argparse.ArgumentParser(description="📄 PDF 工具箱（命令行版）")
    sub = parser.add_subparsers(dest='command', required=True)

    p = sub.add_parser('merge', help='合并多个 PDF')
    p.add_argument('inputs', nargs='+', help='输入 PDF（可多个）')
    p.add_argument('-o', '--output', required=True, help='输出文件')
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser('split', help='拆分 PDF')
    p.add_argument('input', help='输入 PDF')
    p.add_argument('-o', '--output-dir', default='output', help='输出目录')
    p.add_argument('-s', '--size', type=int, default=1, help='每个文件页数（默认 1）')
    p.set_defaults(func=cmd_split)

    p = sub.add_parser('delete', help='删除指定页面')
    p.add_argument('input')
    p.add_argument('-p', '--pages', required=True, help='要删除的页码，如 2,4,6-8')
    p.add_argument('-o', '--output', required=True)
    p.set_defaults(func=cmd_delete)

    p = sub.add_parser('extract', help='提取指定页面')
    p.add_argument('input')
    p.add_argument('-p', '--pages', required=True, help='要提取的页码，如 1,3,5-7')
    p.add_argument('-o', '--output', required=True)
    p.set_defaults(func=cmd_extract)

    p = sub.add_parser('reorder', help='调整页面顺序')
    p.add_argument('input')
    p.add_argument('-r', '--order', required=True, help='新顺序，如 3,1,2')
    p.add_argument('-o', '--output', required=True)
    p.set_defaults(func=cmd_reorder)

    p = sub.add_parser('rotate', help='旋转页面')
    p.add_argument('input')
    p.add_argument('-a', '--angle', type=int, choices=[90, 180, 270, -90],
                   required=True, help='旋转角度（顺时针）')
    p.add_argument('-p', '--pages', help='指定页码，不填则全部')
    p.add_argument('-o', '--output', required=True)
    p.set_defaults(func=cmd_rotate)

    p = sub.add_parser('pdf2img', help='PDF 转图片')
    p.add_argument('input')
    p.add_argument('-o', '--output-dir', default='images')
    p.add_argument('--dpi', type=int, default=150, help='清晰度（默认 150）')
    p.add_argument('--format', default='png', choices=['png', 'jpg'])
    p.set_defaults(func=cmd_pdf2img)

    p = sub.add_parser('img2pdf', help='图片转 PDF')
    p.add_argument('inputs', nargs='+', help='输入图片（可多个）')
    p.add_argument('-o', '--output', required=True)
    p.set_defaults(func=cmd_img2pdf)

    p = sub.add_parser('pdf2word', help='PDF 转 Word')
    p.add_argument('input')
    p.add_argument('-o', '--output', required=True, help='输出 .docx')
    p.set_defaults(func=cmd_pdf2word)

    p = sub.add_parser('office2pdf', help='Word/PPT 转 PDF（需 LibreOffice）')
    p.add_argument('input', help='输入 .docx/.pptx 等')
    p.add_argument('-o', '--output-dir', default='output')
    p.set_defaults(func=cmd_office2pdf)

    p = sub.add_parser('pdf2excel', help='PDF 表格转 Excel')
    p.add_argument('input')
    p.add_argument('-o', '--output', required=True, help='输出 .xlsx')
    p.set_defaults(func=cmd_pdf2excel)

    p = sub.add_parser('pdf2text', help='PDF 转纯文本')
    p.add_argument('input')
    p.add_argument('-o', '--output', required=True, help='输出 .txt')
    p.set_defaults(func=cmd_pdf2text)

    return parser


def main():
    args = build_parser().parse_args()
    try:
        args.func(args)
    except Exception as e:
        print(f"❌ 出错了: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
