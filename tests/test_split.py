#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""不依赖 pytest，纯标准库：python -m unittest discover -s tests -v"""
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from docx import Document  # noqa: E402

import split_doc as sd  # noqa: E402


def make_docx(path: Path, paragraphs: list[str]) -> Path:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(str(path))
    return path


class TestIterSections(unittest.TestCase):
    def test_基本拆分(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = make_docx(Path(tmp) / 'a.docx', ['1.第一题', '2. 第二题', '3.第三题'])
            self.assertEqual(
                list(sd.iter_sections(f)),
                [(1, '第一题'), (2, '第二题'), (3, '第三题')],
            )

    def test_全角句点也认(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = make_docx(Path(tmp) / 'a.docx', ['1．全角句点'])
            self.assertEqual(list(sd.iter_sections(f)), [(1, '全角句点')])

    def test_正文中间的编号不误切(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = make_docx(Path(tmp) / 'a.docx', ['1.第一题正文提到 31. 这个数字', '2.第二题'])
            sections = list(sd.iter_sections(f))
            self.assertEqual(len(sections), 2)
            self.assertIn('31.', sections[0][1])

    def test_没有编号时切不出来(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = make_docx(Path(tmp) / 'a.docx', ['一段没有编号的话'])
            self.assertEqual(list(sd.iter_sections(f)), [])


class TestSplitDocument(unittest.TestCase):
    def test_输出文件数与内容(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = make_docx(tmp / 'src.docx', ['1.甲', '2.乙', '3.丙'])
            out = tmp / 'out'
            count = sd.split_document(src, out, prefix='题目_')

            self.assertEqual(count, 3)
            names = sorted(p.name for p in out.glob('*.docx'))
            self.assertEqual(names, ['题目_1.docx', '题目_2.docx', '题目_3.docx'])
            self.assertEqual(Document(str(out / '题目_2.docx')).paragraphs[0].text, '2.乙')

    def test_无前缀用纯序号(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            src = make_docx(tmp / 'src.docx', ['1.甲'])
            out = tmp / 'out'
            sd.split_document(src, out)
            self.assertTrue((out / '1.docx').exists())


if __name__ == '__main__':
    unittest.main()
