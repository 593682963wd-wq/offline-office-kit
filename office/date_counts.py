"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from datetime import date, datetime, timedelta

def weekday_pattern(ws, row: int) -> str:
    return "".join(
        str(i) for i in range(1, 8)
        if ws.cell(row, 4 + i).value not in (None, "")
    )

def execution_count(ws, row: int, start: date, end: date) -> int:
    cur = start
    n = 0
    while cur <= end:
        if ws.cell(row, 4 + cur.isoweekday()).value not in (None, ""):
            n += 1
        cur += timedelta(days=1)
    return n
