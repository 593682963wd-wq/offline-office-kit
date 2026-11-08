"""作者：Wangyuan；日期：2026-10-29。用显式列名比较两份表格及跨日时段。"""

from __future__ import annotations

import argparse
import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook

from .excel import save_centered_workbook


def parse_clock(value: object) -> int | None:
    """严格解析 HH:MM 或 HHMM（00:00—47:59），空值返回 None。

    扩展小时表示调用方显式给定的次日时间。拒绝从任意文字抽取数字，
    避免把错别字、Excel日期序号或含秒值误当作有效时刻。
    """
    if value is None or value == "":
        return None
    if isinstance(value, dt.time):
        if value.second or value.microsecond or value.tzinfo is not None:
            raise ValueError("仅支持不含秒和时区的时刻")
        return value.hour * 60 + value.minute
    if not isinstance(value, str):
        raise TypeError("时刻须为 HH:MM 文本或 time 对象")
    text = value.strip()
    match = re.fullmatch(r"(\d{1,2}):(\d{2})", text)
    if match:
        hour, minute = map(int, match.groups())
    elif re.fullmatch(r"\d{4}", text):
        hour, minute = int(text[:2]), int(text[2:])
    else:
        raise ValueError("无效时刻格式")
    if hour > 47 or minute > 59:
        raise ValueError("时刻超出 00:00—47:59 范围")
    return hour * 60 + minute


def duration_minutes(start: int | None, end: int | None) -> int | None:
    """计算时长；普通时钟最多跨一次午夜，扩展小时保持显式日序。"""
    if start is None or end is None:
        return None
    if any(
        isinstance(v, bool) or not isinstance(v, int) or not 0 <= v < 2880
        for v in (start, end)
    ):
        raise ValueError("分钟值须在 0—2879 范围")
    if end < start:
        if start >= 1440 or end >= 1440:
            raise ValueError("显式次日时刻顺序不合法")
        end += 1440
    return end - start


def signed_clock_delta(new: int | None, old: int | None) -> int | None:
    """比较时钟平移；普通时钟取最短有向差，扩展小时按绝对分钟。"""
    if new is None or old is None:
        return None
    if new >= 1440 or old >= 1440:
        return new - old
    delta = (new - old) % 1440
    return delta - 1440 if delta > 720 else delta


def format_minutes(value: int | None, *, signed: bool = False) -> str:
    """显示小时:分钟，不把超过一天的时长取模。"""
    if value is None:
        return ""
    sign = "-" if value < 0 else "+" if signed and value > 0 else ""
    hours, minutes = divmod(abs(value), 60)
    return f"{sign}{hours:02d}:{minutes:02d}"


@dataclass(frozen=True)
class TableSchema:
    """完全由调用方指定的表结构，不从文件名或工作目录推断规则。"""

    key_columns: tuple[str, ...]
    compare_columns: tuple[str, ...]
    start_column: str | None = None
    end_column: str | None = None
    header_row: int = 1

    def validate(self) -> None:
        """验证非空主键、列名、时间对及表头行。"""
        if not self.key_columns or not self.compare_columns:
            raise ValueError("请显式指定主键列和比较列")
        if (
            isinstance(self.header_row, bool)
            or not isinstance(self.header_row, int)
            or self.header_row < 1
        ):
            raise ValueError("表头行须为正整数")
        if bool(self.start_column) != bool(self.end_column):
            raise ValueError("开始和结束时刻列必须同时指定")
        if self.start_column and self.start_column == self.end_column:
            raise ValueError("开始和结束时刻列不能相同")
        for group in (self.key_columns, self.compare_columns):
            if any(not isinstance(v, str) or not v.strip() for v in group) or len(
                set(group)
            ) != len(group):
                raise ValueError("列名须为非空且不能重复")
        if self.start_column and (
            self.start_column not in self.compare_columns
            or self.end_column not in self.compare_columns
        ):
            raise ValueError("时刻列必须同时包含在比较列中")


@dataclass(frozen=True)
class Change:
    """一个可追溯到源文件行号的变化，不输出源文件绝对路径。"""

    kind: str
    key: tuple[str, ...]
    field: str
    old: object
    new: object
    old_row: int | None
    new_row: int | None
    delta_minutes: int | None = None


def read_table(
    path: str | Path, schema: TableSchema, *, sheet: str | None = None
) -> dict[tuple[str, ...], tuple[int, dict[str, Any]]]:
    """读取选定列；重复主键、公式及不合法时刻均阻断比较。"""
    schema.validate()
    workbook = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        worksheet = workbook[sheet] if sheet else workbook[workbook.sheetnames[0]]
        header_cells = next(
            worksheet.iter_rows(min_row=schema.header_row, max_row=schema.header_row),
            None,
        )
        if header_cells is None:
            raise ValueError("工作表为空或指定表头行不存在")
        headers = [
            str(c.value).strip() if c.value is not None else "" for c in header_cells
        ]
        columns = tuple(dict.fromkeys(schema.key_columns + schema.compare_columns))
        indexes: dict[str, int] = {}
        for name in columns:
            if headers.count(name) != 1:
                raise ValueError(f"列名不存在或不唯一：{name}")
            indexes[name] = headers.index(name)
        result: dict[tuple[str, ...], tuple[int, dict[str, Any]]] = {}
        for row_number, cells in enumerate(
            worksheet.iter_rows(min_row=schema.header_row + 1), schema.header_row + 1
        ):
            values = {name: cells[index].value for name, index in indexes.items()}
            if all(value in (None, "") for value in values.values()):
                continue
            if any(cells[index].data_type == "f" for index in indexes.values()):
                raise ValueError(f"第 {row_number} 行含公式，请提供明确的值")
            key = tuple(
                str(values[name]).strip() if values[name] is not None else ""
                for name in schema.key_columns
            )
            if any(not value for value in key) or key in result:
                raise ValueError(f"第 {row_number} 行主键为空或重复")
            if schema.start_column and schema.end_column:
                start = parse_clock(values[schema.start_column])
                end = parse_clock(values[schema.end_column])
                duration_minutes(start, end)
            result[key] = (row_number, values)
        return result
    finally:
        workbook.close()


def compare_files(
    old: str | Path, new: str | Path, schema: TableSchema, *, sheet: str | None = None
) -> list[Change]:
    """比较明确旧版与新版；主键变更视为删除/新增，不擅自模糊匹配。"""
    left, right = (
        read_table(old, schema, sheet=sheet),
        read_table(new, schema, sheet=sheet),
    )
    changes: list[Change] = []
    for key in sorted(left.keys() | right.keys()):
        if key not in left:
            row, values = right[key]
            for field in schema.compare_columns:
                changes.append(
                    Change("新增", key, field, None, values[field], None, row)
                )
            continue
        if key not in right:
            row, values = left[key]
            for field in schema.compare_columns:
                changes.append(
                    Change("删除", key, field, values[field], None, row, None)
                )
            continue
        old_row, old_values = left[key]
        new_row, new_values = right[key]
        for field in schema.compare_columns:
            old_value, new_value = old_values[field], new_values[field]
            delta = None
            if field in (schema.start_column, schema.end_column):
                old_value, new_value = parse_clock(old_value), parse_clock(new_value)
                delta = signed_clock_delta(new_value, old_value)
            bool_type_change = isinstance(old_value, bool) != isinstance(new_value, bool)
            if old_value != new_value or bool_type_change:
                changes.append(
                    Change(
                        "调整",
                        key,
                        field,
                        old_values[field],
                        new_values[field],
                        old_row,
                        new_row,
                        delta,
                    )
                )
        if schema.start_column and schema.end_column:
            durations = [
                duration_minutes(
                    parse_clock(v[schema.start_column]),
                    parse_clock(v[schema.end_column]),
                )
                for v in (old_values, new_values)
            ]
            before, after = durations
            if before != after:
                delta = (
                    after - before if before is not None and after is not None else None
                )
                changes.append(
                    Change(
                        "时长变化",
                        key,
                        "计算时长",
                        format_minutes(before),
                        format_minutes(after),
                        old_row,
                        new_row,
                        delta,
                    )
                )
    return changes


def write_changes(changes: list[Change], output: str | Path) -> Path:
    """输出全新差异表；文本明确写为文本，避免外部值成为公式。"""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "差异"
    worksheet.append(["变化", "主键", "字段", "旧值", "新值", "旧行", "新行", "差值"])
    for change in changes:
        worksheet.append(
            [
                change.kind,
                " | ".join(change.key),
                change.field,
                change.old,
                change.new,
                change.old_row,
                change.new_row,
                format_minutes(change.delta_minutes, signed=True),
            ]
        )
        for cell in worksheet[worksheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = "s"
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column in "ABCDEFGH":
        worksheet.column_dimensions[column].width = 24
    return save_centered_workbook(workbook, output)


def main() -> None:
    """命令行仅处理显式参数，不搜索工作目录、不自动联网。"""
    parser = argparse.ArgumentParser(description="离线表格版本比较")
    parser.add_argument("old", type=Path)
    parser.add_argument("new", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--keys", nargs="+", required=True)
    parser.add_argument("--fields", nargs="+", required=True)
    parser.add_argument("--start")
    parser.add_argument("--end")
    parser.add_argument("--sheet")
    parser.add_argument("--header-row", type=int, default=1)
    args = parser.parse_args()
    if args.output.resolve() in {args.old.resolve(), args.new.resolve()}:
        parser.error("输出不能覆盖任一输入文件")
    schema = TableSchema(
        tuple(args.keys), tuple(args.fields), args.start, args.end, args.header_row
    )
    try:
        changes = compare_files(args.old, args.new, schema, sheet=args.sheet)
        write_changes(changes, args.output)
    except (ValueError, TypeError, OSError, KeyError) as error:
        parser.exit(2, f"比较失败：{error}\n")
    print(f"完成：{len(changes)} 条字段级变化")


if __name__ == "__main__":
    main()
