#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 工具箱（图形界面版）

基于 tkinter 的桌面界面，调用 pdf_toolbox.py 中的命令行能力。
支持：合并、拆分、删除页面、提取页面、调整页序、旋转、PDF→图片、图片→PDF、PDF→Word、Office→PDF、PDF→Excel、PDF→文本。
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
CLI_SCRIPT = APP_DIR / "pdf_toolbox.py"


class PdfToolboxGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 工具箱")
        self.geometry("980x680")
        self.minsize(900, 620)

        self.single_input = tk.StringVar()
        self.output_file = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(APP_DIR / "output"))
        self.pages = tk.StringVar()
        self.order = tk.StringVar()
        self.split_size = tk.StringVar(value="1")
        self.rotate_angle = tk.StringVar(value="90")
        self.dpi = tk.StringVar(value="150")
        self.image_format = tk.StringVar(value="png")

        self.multi_inputs = []

        self._build_ui()

    # ---------------- UI ----------------

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(root, text="📄 PDF 工具箱", font=("Microsoft YaHei UI", 18, "bold"))
        title.pack(anchor="w", pady=(0, 10))

        body = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(body, padding=(0, 0, 10, 0))
        right = ttk.Frame(body)
        body.add(left, weight=1)
        body.add(right, weight=1)

        self._build_left(left)
        self._build_right(right)

    def _build_left(self, parent):
        notebook = ttk.Notebook(parent)
        notebook.pack(fill=tk.BOTH, expand=True)

        page_org = ttk.Frame(notebook, padding=10)
        convert = ttk.Frame(notebook, padding=10)
        notebook.add(page_org, text="页面组织")
        notebook.add(convert, text="格式转换")

        self._page_org_tab(page_org)
        self._convert_tab(convert)

    def _build_right(self, parent):
        out_frame = ttk.LabelFrame(parent, text="执行日志", padding=10)
        out_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(out_frame, height=25, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(out_frame, command=self.log_text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scroll.set)

        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_frame, text="清空日志", command=self.clear_log).pack(side=tk.RIGHT)

    def _page_org_tab(self, parent):
        self._input_file_row(parent, "输入 PDF：", self.single_input, [PDF_FILETYPE])
        self._output_file_row(parent, "输出 PDF：", self.output_file, ".pdf")
        self._output_dir_row(parent, "输出目录：", self.output_dir)

        self._entry_row(parent, "页码：", self.pages, "例如：1,3,5-7（删除/提取/旋转可用）")
        self._entry_row(parent, "新顺序：", self.order, "例如：3,1,2（调整页序可用）")
        self._entry_row(parent, "拆分页数：", self.split_size, "每几个页面拆成一个 PDF，默认 1")

        angle_frame = ttk.Frame(parent)
        angle_frame.pack(fill=tk.X, pady=4)
        ttk.Label(angle_frame, text="旋转角度：", width=12).pack(side=tk.LEFT)
        ttk.Combobox(angle_frame, textvariable=self.rotate_angle, values=["90", "180", "270", "-90"], width=10, state="readonly").pack(side=tk.LEFT)

        self._separator(parent)

        ttk.Label(parent, text="多文件输入（合并 PDF 时使用）：").pack(anchor="w")
        self.multi_listbox = tk.Listbox(parent, height=6)
        self.multi_listbox.pack(fill=tk.X, pady=4)

        list_btns = ttk.Frame(parent)
        list_btns.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(list_btns, text="添加 PDF", command=lambda: self.add_multi_files([PDF_FILETYPE])).pack(side=tk.LEFT)
        ttk.Button(list_btns, text="移除选中", command=self.remove_selected_multi).pack(side=tk.LEFT, padx=6)
        ttk.Button(list_btns, text="清空", command=self.clear_multi).pack(side=tk.LEFT)

        action = ttk.LabelFrame(parent, text="操作", padding=8)
        action.pack(fill=tk.X, pady=8)
        buttons = [
            ("合并 PDF", self.run_merge),
            ("拆分 PDF", self.run_split),
            ("删除页面", self.run_delete),
            ("提取页面", self.run_extract),
            ("调整页序", self.run_reorder),
            ("旋转页面", self.run_rotate),
        ]
        self._button_grid(action, buttons)

    def _convert_tab(self, parent):
        self._input_file_row(parent, "输入文件：", self.single_input, [PDF_FILETYPE, IMAGE_FILETYPE, OFFICE_FILETYPE, ALL_FILETYPE])
        self._output_file_row(parent, "输出文件：", self.output_file, "")
        self._output_dir_row(parent, "输出目录：", self.output_dir)

        opt = ttk.LabelFrame(parent, text="转换选项", padding=8)
        opt.pack(fill=tk.X, pady=8)

        row1 = ttk.Frame(opt)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="DPI：", width=12).pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.dpi, width=12).pack(side=tk.LEFT)
        ttk.Label(row1, text="PDF 转图片清晰度，默认 150").pack(side=tk.LEFT, padx=8)

        row2 = ttk.Frame(opt)
        row2.pack(fill=tk.X, pady=4)
        ttk.Label(row2, text="图片格式：", width=12).pack(side=tk.LEFT)
        ttk.Combobox(row2, textvariable=self.image_format, values=["png", "jpg"], width=10, state="readonly").pack(side=tk.LEFT)

        self._separator(parent)

        ttk.Label(parent, text="多图片输入（图片转 PDF 时使用）：").pack(anchor="w")
        self.multi_img_listbox = tk.Listbox(parent, height=6)
        self.multi_img_listbox.pack(fill=tk.X, pady=4)

        list_btns = ttk.Frame(parent)
        list_btns.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(list_btns, text="添加图片", command=lambda: self.add_multi_files([IMAGE_FILETYPE], image_mode=True)).pack(side=tk.LEFT)
        ttk.Button(list_btns, text="移除选中", command=lambda: self.remove_selected_multi(image_mode=True)).pack(side=tk.LEFT, padx=6)
        ttk.Button(list_btns, text="清空", command=lambda: self.clear_multi(image_mode=True)).pack(side=tk.LEFT)

        action = ttk.LabelFrame(parent, text="操作", padding=8)
        action.pack(fill=tk.X, pady=8)
        buttons = [
            ("PDF → 图片", self.run_pdf2img),
            ("图片 → PDF", self.run_img2pdf),
            ("PDF → Word", self.run_pdf2word),
            ("Word/PPT → PDF", self.run_office2pdf),
            ("PDF → Excel", self.run_pdf2excel),
            ("PDF → 纯文本", self.run_pdf2text),
        ]
        self._button_grid(action, buttons)

    def _input_file_row(self, parent, label, variable, filetypes):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text=label, width=12).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=variable).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="选择", command=lambda: self.choose_input_file(variable, filetypes)).pack(side=tk.LEFT, padx=(6, 0))

    def _output_file_row(self, parent, label, variable, default_ext):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text=label, width=12).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=variable).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="选择", command=lambda: self.choose_output_file(variable, default_ext)).pack(side=tk.LEFT, padx=(6, 0))

    def _output_dir_row(self, parent, label, variable):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text=label, width=12).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=variable).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="选择", command=lambda: self.choose_output_dir(variable)).pack(side=tk.LEFT, padx=(6, 0))

    def _entry_row(self, parent, label, variable, hint):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text=label, width=12).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=variable, width=18).pack(side=tk.LEFT)
        ttk.Label(row, text=hint).pack(side=tk.LEFT, padx=8)

    def _button_grid(self, parent, buttons):
        for i, (text, cmd) in enumerate(buttons):
            btn = ttk.Button(parent, text=text, command=cmd)
            btn.grid(row=i // 2, column=i % 2, sticky="ew", padx=4, pady=4)
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

    def _separator(self, parent):
        ttk.Separator(parent).pack(fill=tk.X, pady=10)

    # ---------------- file selection ----------------

    def choose_input_file(self, variable, filetypes):
        path = filedialog.askopenfilename(title="选择输入文件", filetypes=filetypes)
        if path:
            variable.set(path)
            if not self.output_dir.get():
                self.output_dir.set(str(Path(path).parent))

    def choose_output_file(self, variable, default_ext):
        filetypes = [ALL_FILETYPE]
        if default_ext == ".pdf":
            filetypes = [PDF_FILETYPE, ALL_FILETYPE]
        elif default_ext == ".docx":
            filetypes = [("Word 文档", "*.docx"), ALL_FILETYPE]
        elif default_ext == ".xlsx":
            filetypes = [("Excel 文件", "*.xlsx"), ALL_FILETYPE]
        elif default_ext == ".txt":
            filetypes = [("文本文件", "*.txt"), ALL_FILETYPE]

        path = filedialog.asksaveasfilename(title="选择输出文件", defaultextension=default_ext, filetypes=filetypes)
        if path:
            variable.set(path)

    def choose_output_dir(self, variable):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            variable.set(path)

    def add_multi_files(self, filetypes, image_mode=False):
        paths = filedialog.askopenfilenames(title="选择文件", filetypes=filetypes + [ALL_FILETYPE])
        if not paths:
            return
        target = self.multi_img_listbox if image_mode else self.multi_listbox
        for path in paths:
            target.insert(tk.END, path)

    def remove_selected_multi(self, image_mode=False):
        target = self.multi_img_listbox if image_mode else self.multi_listbox
        for idx in reversed(target.curselection()):
            target.delete(idx)

    def clear_multi(self, image_mode=False):
        target = self.multi_img_listbox if image_mode else self.multi_listbox
        target.delete(0, tk.END)

    def get_multi_items(self, image_mode=False):
        target = self.multi_img_listbox if image_mode else self.multi_listbox
        return list(target.get(0, tk.END))

    # ---------------- command runners ----------------

    def run_merge(self):
        inputs = self.get_multi_items()
        if len(inputs) < 2:
            self.error("请至少添加 2 个 PDF 文件用于合并。")
            return
        output = self.require_output_file("请选择输出 PDF 文件。")
        if output:
            self.run_cli(["merge", *inputs, "-o", output])

    def run_split(self):
        input_file = self.require_input_pdf()
        output_dir = self.require_output_dir()
        if input_file and output_dir:
            self.run_cli(["split", input_file, "-s", self.split_size.get().strip() or "1", "-o", output_dir])

    def run_delete(self):
        input_file = self.require_input_pdf()
        output = self.require_output_file("请选择输出 PDF 文件。")
        pages = self.require_pages()
        if input_file and output and pages:
            self.run_cli(["delete", input_file, "-p", pages, "-o", output])

    def run_extract(self):
        input_file = self.require_input_pdf()
        output = self.require_output_file("请选择输出 PDF 文件。")
        pages = self.require_pages()
        if input_file and output and pages:
            self.run_cli(["extract", input_file, "-p", pages, "-o", output])

    def run_reorder(self):
        input_file = self.require_input_pdf()
        output = self.require_output_file("请选择输出 PDF 文件。")
        order = self.order.get().strip()
        if not order:
            self.error("请输入新顺序，例如：3,1,2")
            return
        if input_file and output:
            self.run_cli(["reorder", input_file, "-r", order, "-o", output])

    def run_rotate(self):
        input_file = self.require_input_pdf()
        output = self.require_output_file("请选择输出 PDF 文件。")
        if input_file and output:
            cmd = ["rotate", input_file, "-a", self.rotate_angle.get(), "-o", output]
            pages = self.pages.get().strip()
            if pages:
                cmd.extend(["-p", pages])
            self.run_cli(cmd)

    def run_pdf2img(self):
        input_file = self.require_input_pdf()
        output_dir = self.require_output_dir()
        if input_file and output_dir:
            self.run_cli(["pdf2img", input_file, "-o", output_dir, "--dpi", self.dpi.get().strip() or "150", "--format", self.image_format.get()])

    def run_img2pdf(self):
        inputs = self.get_multi_items(image_mode=True)
        if not inputs:
            self.error("请添加至少 1 张图片。")
            return
        output = self.require_output_file("请选择输出 PDF 文件。")
        if output:
            self.run_cli(["img2pdf", *inputs, "-o", output])

    def run_pdf2word(self):
        input_file = self.require_input_pdf()
        output = self.output_file.get().strip()
        if not output:
            self.error("请选择输出 Word 文件（.docx）。")
            return
        if input_file:
            self.run_cli(["pdf2word", input_file, "-o", output])

    def run_office2pdf(self):
        input_file = self.single_input.get().strip()
        if not input_file:
            self.error("请选择 Word/PPT 输入文件。")
            return
        output_dir = self.require_output_dir()
        if output_dir:
            self.run_cli(["office2pdf", input_file, "-o", output_dir])

    def run_pdf2excel(self):
        input_file = self.require_input_pdf()
        output = self.output_file.get().strip()
        if not output:
            self.error("请选择输出 Excel 文件（.xlsx）。")
            return
        if input_file:
            self.run_cli(["pdf2excel", input_file, "-o", output])

    def run_pdf2text(self):
        input_file = self.require_input_pdf()
        output = self.output_file.get().strip()
        if not output:
            self.error("请选择输出文本文件（.txt）。")
            return
        if input_file:
            self.run_cli(["pdf2text", input_file, "-o", output])

    # ---------------- validation / execution ----------------

    def require_input_pdf(self):
        path = self.single_input.get().strip()
        if not path:
            self.error("请选择输入 PDF 文件。")
            return None
        return path

    def require_output_file(self, msg):
        path = self.output_file.get().strip()
        if not path:
            self.error(msg)
            return None
        return path

    def require_output_dir(self):
        path = self.output_dir.get().strip()
        if not path:
            self.error("请选择输出目录。")
            return None
        return path

    def require_pages(self):
        pages = self.pages.get().strip()
        if not pages:
            self.error("请输入页码，例如：1,3,5-7")
            return None
        return pages

    def run_cli(self, args):
        if not CLI_SCRIPT.exists():
            self.error(f"找不到命令行脚本：{CLI_SCRIPT}")
            return

        command = [sys.executable, str(CLI_SCRIPT), *args]
        self.log("\n$ " + " ".join(f'"{x}"' if " " in x else x for x in command) + "\n")

        def worker():
            try:
                env = os.environ.copy()
                env["PYTHONIOENCODING"] = "utf-8"
                process = subprocess.Popen(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    env=env,
                )
                assert process.stdout is not None
                for line in process.stdout:
                    self.log(line)
                code = process.wait()
                if code == 0:
                    self.log("\n✅ 操作完成\n")
                else:
                    self.log(f"\n❌ 操作失败，退出码：{code}\n")
            except Exception as exc:
                self.log(f"\n❌ 执行出错：{exc}\n")

        threading.Thread(target=worker, daemon=True).start()

    def log(self, message):
        self.log_text.after(0, lambda: self._append_log(message))

    def _append_log(self, message):
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)

    def clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def error(self, message):
        messagebox.showerror("提示", message)


PDF_FILETYPE = ("PDF 文件", "*.pdf")
IMAGE_FILETYPE = ("图片文件", "*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp")
OFFICE_FILETYPE = ("Office 文件", "*.doc;*.docx;*.ppt;*.pptx")
ALL_FILETYPE = ("所有文件", "*.*")


def main():
    app = PdfToolboxGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
