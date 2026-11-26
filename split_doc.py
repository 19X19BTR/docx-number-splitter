#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档按行首序号拆分：`1.` `2.` `31.` 这种编号即切口，一键拆成多个 docx。

特点：
- 支持「点后无空格」的写法（1.正文），也能避免正文里出现的 31. 被误切
- 支持中文全角句点（1．正文）

用法：
    python split_doc.py                            # 弹窗选文件
    python split_doc.py --input 素材.docx --outdir ./split_result --prefix 题目_
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

from docx import Document

# 行首 1-3 位数字 + 半角/全角点，点后可有可无空格
NUMBER_PATTERN = re.compile(r'(?:^|\n)(\d{1,3})(?:[.\uff0e])(.*?)(?=\n\d{1,3}[.\uff0e]|$)', re.S)


def pick_file(title: str = '请选择要拆分的 Word 文档') -> str:
    from tkinter import Tk, filedialog

    Tk().withdraw()
    return filedialog.askopenfilename(title=title, filetypes=[('Word 文档', '*.docx')])


def pick_dir(title: str = '请选择输出目录') -> str:
    from tkinter import Tk, filedialog

    Tk().withdraw()
    return filedialog.askdirectory(title=title)


def iter_sections(docx_path: Path):
    """按编号切口切出 (序号, 正文) 序列。"""
    doc = Document(str(docx_path))
    full_text = re.sub(r'\n+', '\n', '\n'.join(p.text for p in doc.paragraphs))
    for match in NUMBER_PATTERN.finditer(full_text):
        yield int(match.group(1)), match.group(2).strip()


def split_document(docx_path: Path, out_dir: Path, prefix: str = '') -> int:
    out_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    for index, body in iter_sections(docx_path):
        new_doc = Document()
        new_doc.add_paragraph(f'{index}.{body}')
        name = f'{prefix}{index}.docx' if prefix else f'{index}.docx'
        new_doc.save(str(out_dir / name))
        count += 1
    return count


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description='Word 文档按序号拆分')
    parser.add_argument('--input', help='要拆分的 .docx')
    parser.add_argument('--outdir', help='输出目录，默认 <源文件目录>/split_result')
    parser.add_argument('--prefix', default='', help='输出文件名前缀，留空则用纯序号')
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)

    src = Path(args.input).expanduser().resolve() if args.input else Path(pick_file())
    if not src.exists():
        print('未选择文件，程序退出。')
        return 1

    out_dir = Path(args.outdir).expanduser().resolve() if args.outdir else src.parent / 'split_result'
    if not args.outdir and not args.input:
        chosen = pick_dir()
        if chosen:
            out_dir = Path(chosen)

    prefix = args.prefix
    if not args.input:
        prefix = input('请输入文件名前缀（回车=纯序号）：').strip()

    count = split_document(src, out_dir, prefix)
    print(f'全部拆完！共 {count} 个文件 → {out_dir}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
