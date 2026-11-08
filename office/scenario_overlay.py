"""作者：Wangyuan；日期：2026-10-29。显式增量表的时长场景叠加。"""
from __future__ import annotations
import math
from typing import Mapping, Any
from bisect import bisect_right


def overlay(rows: list[dict[str,Any]], *, increments: Mapping[tuple[str,str],float], boundaries: tuple[float,...]) -> list[dict[str,Any]]:
    """逐航段计算基线缺口和增量缺口；缺少增量时保持未知，不估算或套用运营等级。"""
    if any(not math.isfinite(v) for v in boundaries) or tuple(sorted(set(boundaries)))!=boundaries:raise ValueError("分档边界须有限严格升序")
    output=[]
    for row in rows:
        planned=row.get("planned_minutes");history=row.get("history_minutes");key=(row["origin"],row["destination"]);delta=increments.get(key)
        if any(v is not None and (isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v)) for v in (planned,history,delta)):raise ValueError("时长与增量须为有限数")
        if any(v is not None and v<0 for v in (planned,history)):raise ValueError("基线时长不能为负")
        before=history-planned if history is not None and planned is not None else None
        after=before+delta if before is not None and delta is not None else None
        output.append({**row,"increment_minutes":delta,"base_gap":before,"scenario_gap":after,"base_band":bisect_right(boundaries,before) if before is not None else None,"scenario_band":bisect_right(boundaries,after) if after is not None else None,"status":"complete" if after is not None else "missing_input"})
    return output
