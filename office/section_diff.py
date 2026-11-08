"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from dataclasses import dataclass, field

import difflib

import re

import unicodedata

_NOISE_RE = re.compile(r"^(?:Page\s+\d+|<!--\s*page:\d+\s*-->)$", re.I)

_SECTION_RE = re.compile(r"^Section\s+(\d+(?:\.\d+)?)(?:\s+(.+))?$", re.I)

@dataclass(frozen=True)
class SectionDiffItem:
    section: str
    change_type: str
    old_text: str = ""
    new_text: str = ""

def _clean_compare_line(line: str) -> str:
    line = re.sub(r"\s+", " ", line).strip()
    if not line or _NOISE_RE.match(line):
        return ""
    return line

def markdown_sections(markdown: str) -> dict[str, list[str]]:
    sections: dict[str, list[str]] = {"未分类": []}
    current = "未分类"
    for raw_line in markdown.splitlines():
        line = _clean_compare_line(raw_line.lstrip("# "))
        if not line:
            continue
        match = _SECTION_RE.search(line)
        if match:
            suffix = (match.group(2) or "").strip()
            current = f"Section {match.group(1)} {suffix}".rstrip()
            sections.setdefault(current, []).append(line)
            continue
        sections.setdefault(current, []).append(line)
    return {name: lines for name, lines in sections.items() if lines}

def _normalise_for_compare(text: str) -> str:
    # Equality suppression must never erase decimal points, runway separators,
    # coordinate signs, frequency punctuation or other business-bearing marks.
    # NFKC only removes representation differences such as full-width glyphs;
    # whitespace remains the sole ignored layout noise.
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text)).casefold()

def compare_markdown(old_markdown: str, new_markdown: str) -> list[SectionDiffItem]:
    old_sections = markdown_sections(old_markdown)
    new_sections = markdown_sections(new_markdown)
    section_order = list(dict.fromkeys([*old_sections, *new_sections]))
    items: list[SectionDiffItem] = []
    for section in section_order:
        old_lines = old_sections.get(section, [])
        new_lines = new_sections.get(section, [])
        matcher = difflib.SequenceMatcher(a=old_lines, b=new_lines, autojunk=False)
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                continue
            old_block, new_block = old_lines[i1:i2], new_lines[j1:j2]
            if tag == "replace":
                count = max(len(old_block), len(new_block))
                for index in range(count):
                    old_text = old_block[index] if index < len(old_block) else ""
                    new_text = new_block[index] if index < len(new_block) else ""
                    if old_text and new_text and _normalise_for_compare(old_text) == _normalise_for_compare(new_text):
                        continue
                    change_type = "修改" if old_text and new_text else "新增" if new_text else "删除"
                    items.append(SectionDiffItem(section, change_type, old_text, new_text))
            elif tag == "delete":
                items.extend(SectionDiffItem(section, "删除", old_text, "") for old_text in old_block)
            elif tag == "insert":
                items.extend(SectionDiffItem(section, "新增", "", new_text) for new_text in new_block)
    return items
