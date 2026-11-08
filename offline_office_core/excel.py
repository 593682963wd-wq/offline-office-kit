"""作者：Wangyuan；日期：2026-10-29。仅修改显式传入工作簿的对齐属性。"""

from __future__ import annotations

import os
import tempfile
from copy import copy
from pathlib import Path

from openpyxl.cell.cell import MergedCell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def center_worksheet(worksheet: Worksheet) -> Worksheet:
    """居中有内容或格式的单元格，保留公式、换行、字体和边框。"""
    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value not in (None, "") or cell.has_style:
                alignment = copy(cell.alignment)
                alignment.horizontal = "center"
                alignment.vertical = "center"
                cell.alignment = alignment
    return worksheet


def center_workbook(workbook: Workbook) -> Workbook:
    """居中普通工作簿；写入专用模式无法回读单元格，明确拒绝。"""
    if workbook.write_only or workbook.read_only:
        raise ValueError("请传入普通读写模式的工作簿")
    for worksheet in workbook.worksheets:
        center_worksheet(worksheet)
    return workbook


def save_centered_workbook(workbook: Workbook, target: str | Path) -> Path:
    """原子保存显式目标文件，不自动创建目录或查找其他工作簿。"""
    destination = Path(target).resolve()
    if not destination.parent.is_dir():
        raise FileNotFoundError(destination.parent)
    center_workbook(workbook)
    fd, name = tempfile.mkstemp(
        prefix=".office-", suffix=".xlsx", dir=destination.parent
    )
    os.close(fd)
    temporary = Path(name)
    try:
        workbook.save(temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)
    return destination
