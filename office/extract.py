"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from datetime import date, datetime, time

from pathlib import Path

from typing import Any, Callable

def _openpyxl_to_text(path: Path, raw: dict[str, Any]) -> str:
    try:
        from openpyxl import load_workbook
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError("读取 Excel 需要 openpyxl，请在运行环境安装 openpyxl。") from exc

    sheet_name = str(raw.get("sheet") or "").strip()
    max_rows = int(raw.get("max_rows") or 300)
    max_cols = int(raw.get("max_cols") or 40)
    workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        worksheet = workbook[sheet_name] if sheet_name else workbook[workbook.sheetnames[0]]
        lines: list[str] = []
        for row in worksheet.iter_rows(max_row=max_rows, max_col=max_cols, values_only=True):
            values = [_cell_to_text(value) for value in row]
            while values and not values[-1]:
                values.pop()
            if any(values):
                lines.append("\t".join(values))
        return "\n".join(lines)
    finally:
        workbook.close()

def _cell_to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()
