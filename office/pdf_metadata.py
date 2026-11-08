"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from dataclasses import dataclass, field

from datetime import date

from pathlib import Path

from typing import Iterable, Sequence

import re

_ICAO_RE = re.compile(r"\b(Z[A-Z]{3})(?:/[A-Z]{3})?\b")

_EFF_RE = re.compile(r"\bEFF\s*[:：]?\s*(\d{4}[-./]\d{1,2}[-./]\d{1,2})", re.I)

_DATE_RE = re.compile(r"\b(\d{4}[-./]\d{1,2}[-./]\d{1,2})\b")

@dataclass(frozen=True)
class PdfMetadata:
    icao: str = "UNKNOWN"
    eff: str = "UNKNOWN"
    data_date: str = "UNKNOWN"
    airport_name: str = "机场名称待确认"

    @property
    def eff_date(self) -> date | None:
        if self.eff == "UNKNOWN":
            return None
        try:
            year, month, day = (int(value) for value in self.eff.split("-"))
            return date(year, month, day)
        except (TypeError, ValueError):
            return None

class PairingError(ValueError):
    """Raised when every uploaded old/new PDF cannot be paired safely."""

def _normalise_date(value: str) -> str:
    parts = re.split(r"[-./]", value)
    if len(parts) != 3:
        return "UNKNOWN"
    try:
        return date(*map(int,parts)).isoformat()
    except ValueError:
        return "UNKNOWN"

def _clean_airport_name(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip(" -—:：/")
    text = re.split(r"\b(?:AD|EFF|NAIP|CAAC)\b", text, maxsplit=1, flags=re.I)[0].strip()
    chinese = re.search(r"([\u4e00-\u9fff][\u4e00-\u9fff·（）()\s]{1,30}(?:机场|航空港))", text)
    return chinese.group(1).strip() if chinese else "机场名称待确认"

def extract_metadata(pdf_path: str | Path) -> PdfMetadata:
    """Read metadata from actual PDF content; filenames are never trusted."""
    import pymupdf

    path = Path(pdf_path)
    with pymupdf.open(path) as document:
        if document.page_count == 0:
            return PdfMetadata()
        first_text = "\n".join(document[index].get_text("text") for index in range(min(2, document.page_count)))

    icao_match = _ICAO_RE.search(first_text.upper())
    eff_match = _EFF_RE.search(first_text)
    all_dates = [_normalise_date(match) for match in _DATE_RE.findall(first_text)]
    eff = _normalise_date(eff_match.group(1)) if eff_match else "UNKNOWN"
    data_date = next((value for value in all_dates if value != eff), "UNKNOWN")

    airport_name = "机场名称待确认"
    if icao_match:
        icao = icao_match.group(1)
        for line in first_text.splitlines():
            if icao in line.upper():
                candidate = re.sub(rf".*?{re.escape(icao)}(?:/[A-Z]{{3}})?\s*[-—:]?", "", line, flags=re.I)
                airport_name = _clean_airport_name(candidate)
                if airport_name != "机场名称待确认":
                    break
    else:
        icao = "UNKNOWN"
    if airport_name == "机场名称待确认":
        airport_name = _clean_airport_name(first_text)
    return PdfMetadata(icao=icao, eff=eff, data_date=data_date, airport_name=airport_name)

def pair_pdf_paths(old_paths: Sequence[str | Path], new_paths: Sequence[str | Path]) -> list[tuple[str, Path, Path]]:
    """Pair by metadata ICAO first; 拒绝依赖上传顺序推断配对."""
    old_entries = [(Path(path), extract_metadata(path)) for path in old_paths]
    new_entries = [(Path(path), extract_metadata(path)) for path in new_paths]
    pairs: list[tuple[str, Path, Path]] = []
    used_old: set[Path] = set()
    used_new: set[Path] = set()
    common = sorted(
        {meta.icao for _, meta in old_entries if meta.icao != "UNKNOWN"}
        & {meta.icao for _, meta in new_entries if meta.icao != "UNKNOWN"}
    )
    for icao in common:
        matching_old = [path for path, meta in old_entries if meta.icao == icao and path not in used_old]
        matching_new = [path for path, meta in new_entries if meta.icao == icao and path not in used_new]
        if len(matching_old)!=1 or len(matching_new)!=1:raise PairingError("元数据重复，不能自动配对")
        for old_path, new_path in zip(matching_old, matching_new):
            pairs.append((icao, old_path, new_path))
            used_old.add(old_path)
            used_new.add(new_path)
    remaining_old = [entry for entry in old_entries if entry[0] not in used_old]
    remaining_new = [entry for entry in new_entries if entry[0] not in used_new]
    if remaining_old or remaining_new:
        raise PairingError("未能按唯一元数据完成配对；需要调用方明确新旧对应")
    if len(pairs) != len(old_entries) or len(pairs) != len(new_entries):
        raise PairingError("新旧 PDF 数量或机场对应关系不完整，已停止生成结果。")
    return pairs
