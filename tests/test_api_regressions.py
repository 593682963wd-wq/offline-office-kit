"""作者：Wangyuan；日期：2026-11-08。合成办公文件类型、空段及差异回归。"""
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches
from office import ppt_edit, pdf_text, section_diff
from offline_office_core.compare import TableSchema, compare_files


class ApiRegressionTests(unittest.TestCase):
    def test_boolean_and_integer_are_different_cells(self):
        with tempfile.TemporaryDirectory() as folder:
            paths = []
            for i, value in enumerate((True, 1)):
                workbook = Workbook(); workbook.active.append(['编号', '值']); workbook.active.append(['same', value])
                path = Path(folder) / f'{i}.xlsx'; workbook.save(path); workbook.close(); paths.append(path)
            self.assertEqual(len(compare_files(*paths, TableSchema(('编号',), ('值',)))), 1)

    def test_clear_all_paragraphs_stays_editable(self):
        prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
        tf = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1)).text_frame
        tf.text = '合成旧字'
        ppt_edit.set_multi_paragraphs(tf, lambda _: False, [])
        self.assertEqual(len(tf.paragraphs), 1)
        ppt_edit.set_text_keep_format(tf, '合成新字')
        self.assertEqual(tf.text, '合成新字')

    def test_table_pipe_is_literal_cell_text(self):
        rows = pdf_text._table_to_markdown([['A', 'B'], ['x|y', 'z']])
        self.assertIn(r'x\|y', rows[-1])

    def test_header_only_change_and_addition_remain_visible(self):
        self.assertTrue(section_diff.compare_markdown('Section 1 限高100', 'Section 1 限高200'))
        self.assertTrue(section_diff.compare_markdown('', 'Section 2 合成标题'))
        self.assertFalse(section_diff.compare_markdown('Section 1 相同', 'Section 1 相同'))


if __name__ == '__main__':
    unittest.main()
