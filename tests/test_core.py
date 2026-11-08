"""作者：Wangyuan；日期：2026-10-29。完全合成的离线行为测试。"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from offline_office_core.compare import (
    Change,
    TableSchema,
    compare_files,
    duration_minutes,
    format_minutes,
    parse_clock,
    read_table,
    signed_clock_delta,
    write_changes,
)
from offline_office_core.docx_table import fill_table
from offline_office_core.excel import center_workbook, save_centered_workbook
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font


class ClockTests(unittest.TestCase):
    def test_midnight_and_explicit_day(self) -> None:
        self.assertEqual(
            duration_minutes(parse_clock("23:30"), parse_clock("01:00")), 90
        )
        self.assertEqual(
            duration_minutes(parse_clock("23:30"), parse_clock("25:00")), 90
        )
        self.assertEqual(parse_clock(dt.time(9, 15)), 555)
        self.assertEqual(parse_clock("0915"), 555)
        self.assertEqual(parse_clock("47:59"), 2879)
        self.assertEqual(duration_minutes(0, 0), 0)

    def test_invalid_inputs_are_rejected(self) -> None:
        for value in (
            "25:90",
            "48:00",
            "x12:30",
            "12:3",
            "12:30:01",
            "-0130",
            1230,
            True,
            dt.time(9, 0, 1),
        ):
            with self.subTest(value=value), self.assertRaises((ValueError, TypeError)):
                parse_clock(value)
        with self.assertRaises(ValueError):
            duration_minutes(1500, 60)
        with self.assertRaises(ValueError):
            duration_minutes(-1, 30)

    def test_missing_is_not_zero(self) -> None:
        self.assertIsNone(parse_clock(None))
        self.assertIsNone(duration_minutes(None, 90))
        self.assertIsNone(signed_clock_delta(None, 0))
        self.assertEqual(format_minutes(None), "")
        self.assertEqual(format_minutes(0), "00:00")
        self.assertEqual(format_minutes(1500), "25:00")
        self.assertEqual(signed_clock_delta(15, 1425), 30)
        self.assertEqual(signed_clock_delta(1425, 15), -30)


class FileTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.schema = TableSchema(("编号",), ("说明", "开始", "结束"), "开始", "结束")

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def make_sheet(
        self, name: str, rows: list[list[object]], headers: list[str] | None = None
    ) -> Path:
        workbook = Workbook()
        workbook.active.append(headers or ["编号", "说明", "开始", "结束"])
        for row in rows:
            workbook.active.append(row)
        path = self.root / name
        workbook.save(path)
        return path

    def test_comparison_round_trip(self) -> None:
        old = self.make_sheet(
            "old.xlsx",
            [
                ["demo-1", "虚构任务甲", "23:30", "01:00"],
                ["demo-2", "虚构任务乙", "09:00", "10:00"],
            ],
        )
        new = self.make_sheet(
            "new.xlsx",
            [
                ["demo-1", "虚构任务甲", "23:30", "01:20"],
                ["demo-3", "虚构任务丙", "11:00", "12:00"],
            ],
        )
        changes = compare_files(old, new, self.schema)
        self.assertEqual(len(changes), 8)
        duration = [c for c in changes if c.kind == "时长变化"]
        self.assertEqual(
            (duration[0].old, duration[0].new, duration[0].delta_minutes),
            ("01:30", "01:50", 20),
        )
        output = write_changes(changes, self.root / "result.xlsx")
        workbook = load_workbook(output)
        self.assertEqual(workbook.active.max_row, 9)
        self.assertEqual(workbook.active["A2"].alignment.horizontal, "center")
        workbook.close()

    def test_duplicate_and_missing_key(self) -> None:
        for rows in (
            [["same", "合成", "09:00", "10:00"], ["same", "合成", "11:00", "12:00"]],
            [[None, "合成", "09:00", "10:00"]],
        ):
            path = self.make_sheet("invalid.xlsx", rows)
            with self.assertRaises(ValueError):
                read_table(path, self.schema)

    def test_missing_and_duplicate_columns(self) -> None:
        for headers in (
            ["编号", "说明", "开始", "备注"],
            ["编号", "说明", "开始", "结束", "编号"],
        ):
            path = self.make_sheet("headers.xlsx", [], headers)
            with self.assertRaises(ValueError):
                read_table(path, self.schema)

    def test_formulas_rejected_and_export_text_safe(self) -> None:
        path = self.make_sheet("formula.xlsx", [["demo", "=1+2", "09:00", "10:00"]])
        with self.assertRaises(ValueError):
            read_table(path, self.schema)
        output = write_changes(
            [Change("调整", ("=identity",), "说明", "=1+2", "+sum", 2, 2)],
            self.root / "literal.xlsx",
        )
        workbook = load_workbook(output, data_only=False)
        for address in ("B2", "D2", "E2"):
            self.assertEqual(workbook.active[address].data_type, "s")
        self.assertEqual(workbook.active["D2"].value, "=1+2")
        workbook.close()

    def test_no_false_change_for_same_time(self) -> None:
        old = self.make_sheet("old.xlsx", [["demo", "合成", "09:00", "10:00"]])
        new = self.make_sheet("new.xlsx", [["demo", "合成", "0900", "1000"]])
        self.assertEqual(compare_files(old, new, self.schema), [])

    def test_empty_tables(self) -> None:
        old = self.make_sheet("old.xlsx", [])
        new = self.make_sheet("new.xlsx", [])
        self.assertEqual(compare_files(old, new, self.schema), [])
        output = write_changes([], self.root / "empty.xlsx")
        workbook = load_workbook(output)
        self.assertEqual(workbook.active.max_row, 1)
        workbook.close()

    def test_fully_empty_workbook_has_friendly_error(self) -> None:
        blank = self.root / "blank.xlsx"
        Workbook().save(blank)
        with self.assertRaisesRegex(ValueError, "工作表为空"):
            read_table(blank, self.schema)
        process = subprocess.run(
            [
                sys.executable,
                "-m",
                "offline_office_core.compare",
                str(blank),
                str(blank),
                str(self.root / "output.xlsx"),
                "--keys",
                "编号",
                "--fields",
                "说明",
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(process.returncode, 2)
        self.assertIn("工作表为空", process.stderr)
        self.assertNotIn("Traceback", process.stderr)
        self.assertFalse((self.root / "output.xlsx").exists())

    def test_alignment_preserves_formula_and_style(self) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "=1+2"
        sheet["A1"].font = Font(bold=True, color="FF0000")
        sheet["A1"].alignment = Alignment(wrap_text=True, indent=1)
        sheet.merge_cells("B1:C1")
        sheet["B1"] = "合成合并标题"
        path = save_centered_workbook(workbook, self.root / "centered.xlsx")
        loaded = load_workbook(path)
        cell = loaded.active["A1"]
        self.assertEqual(cell.value, "=1+2")
        self.assertTrue(cell.font.bold)
        self.assertTrue(cell.alignment.wrap_text)
        self.assertEqual(cell.alignment.indent, 1)
        self.assertEqual(cell.alignment.horizontal, "center")
        self.assertEqual(cell.alignment.vertical, "center")
        self.assertIn("B1:C1", loaded.active.merged_cells)
        loaded.close()
        with self.assertRaises(ValueError):
            center_workbook(Workbook(write_only=True))


class DocumentTests(unittest.TestCase):
    def make_template(self, copies: int = 1) -> BytesIO:
        document = Document()
        document.add_paragraph("虚构文档，所有数据均为合成。")
        for _ in range(copies):
            table = document.add_table(rows=2, cols=2)
            table.rows[0].cells[0].text = "编号"
            table.rows[0].cells[1].text = "说明"
            table.rows[1].cells[0].text = "old"
            table.rows[1].cells[1].text = "旧合成内容"
        output = BytesIO()
        document.save(output)
        output.seek(0)
        return output

    def test_replace_and_append(self) -> None:
        for mode, expected_rows in (("replace", 2), ("append", 3)):
            result = fill_table(
                self.make_template(),
                headers=["编号", "说明"],
                rows=[["new", "新合成内容"]],
                mode=mode,
            )
            table = Document(result).tables[0]
            self.assertEqual(len(table.rows), expected_rows)
            self.assertEqual(table.rows[-1].cells[1].text, "新合成内容")

    def test_missing_ambiguous_and_misaligned_input(self) -> None:
        for source, headers, rows in (
            (self.make_template(), ["不存在"], [["x"]]),
            (self.make_template(2), ["编号", "说明"], [["x", "y"]]),
            (self.make_template(), ["编号", "说明"], [["x"]]),
        ):
            with self.assertRaises(ValueError):
                fill_table(source, headers=headers, rows=rows)

    def test_empty_replacement_preserves_header(self) -> None:
        result = fill_table(self.make_template(), headers=["编号", "说明"], rows=[])
        table = Document(result).tables[0]
        self.assertEqual(len(table.rows), 1)
        self.assertEqual(table.rows[0].cells[0].text, "编号")

    def test_nested_tables_are_rejected_before_replacement_or_append(self) -> None:
        for mode in ("replace", "append"):
            document = Document(self.make_template())
            nested = document.tables[0].rows[1].cells[0].add_table(rows=1, cols=1)
            nested.cell(0, 0).text = "旧合成嵌套内容"
            source = BytesIO()
            document.save(source)
            original = source.getvalue()
            source.seek(0)
            with self.assertRaisesRegex(ValueError, "嵌套表"):
                fill_table(
                    source,
                    headers=["编号", "说明"],
                    rows=[["new", "新合成内容"]],
                    mode=mode,
                )
            self.assertEqual(source.getvalue(), original)

    def test_cell_content_controls_are_rejected(self) -> None:
        document = Document(self.make_template())
        control = OxmlElement("w:sdt")
        document.tables[0].rows[1].cells[0]._tc.append(control)
        source = BytesIO()
        document.save(source)
        source.seek(0)
        with self.assertRaisesRegex(ValueError, "内容控件"):
            fill_table(source, headers=["编号", "说明"], rows=[["new", "合成内容"]])


if __name__ == "__main__":
    unittest.main()
