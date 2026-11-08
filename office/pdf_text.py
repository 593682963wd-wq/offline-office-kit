"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from pathlib import Path

from typing import Iterable, Sequence

import re

def _format_text_line(line):
    return re.sub(r"\s+", " ", line).strip()

def _inside_table(block_bbox: Sequence[float], table_bbox: Sequence[float]) -> bool:
    x0, y0, x1, y1 = block_bbox
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    tx0, ty0, tx1, ty1 = table_bbox
    return tx0 - 3 <= cx <= tx1 + 3 and ty0 - 3 <= cy <= ty1 + 3

def _table_to_markdown(data: Sequence[Sequence[object | None]]) -> list[str]:
    if not data:
        return []
    width = max((len(row) for row in data), default=0)
    if width == 0:
        return []
    cleaned: list[list[str]] = []
    for row in data:
        values = [re.sub(r"\s*\n\s*", "<br>", str(cell)).strip().replace('\\', '\\\\').replace('|', r'\|') if cell is not None else "" for cell in row]
        cleaned.append(values + [""] * (width - len(values)))
    lines = ["| " + " | ".join(cleaned[0]) + " |", "| " + " | ".join(["---"] * width) + " |"]
    lines.extend("| " + " | ".join(row) + " |" for row in cleaned[1:])
    return lines

def _pdf_to_markdown_with_warnings(pdf_path: str | Path) -> tuple[str, list[str]]:
    """Convert PDF to Markdown and report table-extraction degradation."""
    import pymupdf

    pages: list[str] = []
    warnings: list[str] = []
    with pymupdf.open(Path(pdf_path)) as document:
        for page_index, page in enumerate(document, 1):
            page_lines: list[str] = [f"<!-- page:{page_index} -->"]
            try:
                finder = page.find_tables()
                tables = list(finder.tables) if finder else []
            except Exception as exc:
                tables = []
                warnings.append(f"第 {page_index} 页表格识别失败，已按普通文本提取：{type(exc).__name__}")
            table_bboxes = [table.bbox for table in tables]
            blocks = page.get_text("dict").get("blocks", [])
            for block in blocks:
                if block.get("type") != 0 or any(_inside_table(block.get("bbox", (0, 0, 0, 0)), bbox) for bbox in table_bboxes):
                    continue
                for line in block.get("lines", []):
                    text = " ".join(span.get("text", "") for span in line.get("spans", []))
                    formatted = _format_text_line(text)
                    if formatted:
                        page_lines.append(formatted)
            for table in tables:
                table_lines = _table_to_markdown(table.extract() or [])
                if table_lines:
                    page_lines.extend(["", *table_lines, ""])
            if not tables and len(page_lines) == 1:
                for raw_line in page.get_text("text").splitlines():
                    formatted = _format_text_line(raw_line)
                    if formatted:
                        page_lines.append(formatted)
            pages.append("\n".join(page_lines))
    return "\n\n---\n\n".join(pages), warnings
