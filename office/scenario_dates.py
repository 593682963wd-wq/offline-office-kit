"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from datetime import date, datetime, timedelta

def _exec_in_window(weekday_str: str, start: date, end: date) -> int:
    days = {int(c) for c in (weekday_str or "") if c.isdigit()}
    if not days:
        return 0
    cur, n = start, 0
    while cur <= end:
        if cur.isoweekday() in days:
            n += 1
        cur += timedelta(days=1)
    return n
