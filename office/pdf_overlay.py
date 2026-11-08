"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from pathlib import Path

from typing import Any

import json

import math

import re

import fitz

_LATIN_TEXT_RE = re.compile(r"[A-Za-z]")

def _normalize_text(value: str) -> str:
    """合并 PDF 文本块中的换行与多余空白。"""
    return " ".join(str(value or "").replace("\t", " ").split())

def _load_base_translation(path: Path | None) -> dict[str, Any]:
    """读取既有翻译模板；未提供时返回稳定的默认结构。"""
    if path is None:
        return {"sourceLanguage": "en", "targetLanguage": "zh-Hans", "entries": []}
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise TypeError("已有翻译 JSON 的顶层必须是对象。")
    entries = payload.get("entries", [])
    if not isinstance(entries, list):
        raise TypeError("已有翻译 JSON 的 entries 必须是数组。")
    return payload

def generate_translation_template(
    input_pdf: Path,
    template_json: Path,
    base_translation_json: Path | None = None,
) -> dict[str, Any]:
    """从可检索 PDF 提取英文文本块并生成翻译模板。"""
    base = _load_base_translation(base_translation_json)
    existing = {
        str(item.get("sourceText") or ""): str(item.get("translation") or "")
        for item in base.get("entries", [])
        if isinstance(item, dict) and str(item.get("sourceText") or "").strip()
    }
    entries: list[dict[str, Any]] = []
    with fitz.open(input_pdf) as document:
        for page_index, page in enumerate(document):
            for block in page.get_text("blocks", sort=True):
                if len(block) < 7 or int(block[6]) != 0:
                    continue
                source_text = _normalize_text(str(block[4] or ""))
                if not source_text or not _LATIN_TEXT_RE.search(source_text):
                    continue
                entries.append(
                    {
                        "page": page_index + 1,
                        "sourceText": source_text,
                        "translation": existing.get(source_text, ""),
                    }
                )

    payload = {
        "sourceLanguage": str(base.get("sourceLanguage") or "en"),
        "targetLanguage": str(base.get("targetLanguage") or "zh-Hans"),
        "entries": entries,
    }
    template_json.parent.mkdir(parents=True, exist_ok=True)
    template_json.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return payload

def _translation_font_size(
    translations: list[str],
    panel_width: float,
    available_height: float,
) -> float:
    """按当前页译文总量估算不会明显溢出的字号。"""
    for font_size in (7.5, 7.0, 6.5, 6.0, 5.5, 5.0, 4.5, 4.0):
        chars_per_line = max(8, int((panel_width - 20) / font_size))
        required = sum(
            max(1, math.ceil(len(text) / chars_per_line)) * font_size * 1.45 + 6
            for text in translations
        )
        if required <= available_height:
            return font_size
    return 4.0

def _translation_required_height(
    translations: list[str],
    panel_width: float,
    font_size: float,
) -> float:
    """保守估算译文侧栏高度，避免长文在页尾被静默截断。"""
    chars_per_line = max(8, int((panel_width - 20) / font_size))
    return sum(
        max(1, math.ceil(len(text) / chars_per_line)) * font_size * 1.45 + 11
        for text in translations
    )

def render_bilingual_pdf(
    input_pdf: Path,
    output_pdf: Path,
    translation_json: Path,
) -> None:
    """保留原页并在右侧扩展区域写入对应页的中文译文。"""
    payload = json.loads(translation_json.read_text(encoding="utf-8"))
    raw_entries = payload.get("entries", [])
    if not isinstance(raw_entries, list):
        raise TypeError("翻译 JSON 的 entries 必须是数组。")

    translations_by_page: dict[int, list[str]] = {}
    for item in raw_entries:
        if not isinstance(item, dict):
            continue
        translation = _normalize_text(str(item.get("translation") or ""))
        if not translation:
            continue
        try:
            page_number = int(item.get("page"))
        except (TypeError, ValueError):
            continue
        translations_by_page.setdefault(page_number, []).append(translation)

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    output = fitz.open()
    try:
        with fitz.open(input_pdf) as source:
            for page_index, source_page in enumerate(source):
                source_rect = source_page.rect
                panel_width = min(max(source_rect.width * 0.72, 280), 420)
                translations = translations_by_page.get(page_index + 1, [])
                panel_text_width = panel_width - 32
                font_size = _translation_font_size(
                    translations,
                    panel_text_width,
                    source_rect.height - 36,
                )
                required_height = _translation_required_height(
                    translations,
                    panel_text_width,
                    font_size,
                )
                output_height = max(source_rect.height, required_height + 36)
                output_page = output.new_page(
                    width=source_rect.width + panel_width,
                    height=output_height,
                )
                output_page.show_pdf_page(source_rect, source, page_index)
                separator_x = source_rect.width + 10
                output_page.draw_line(
                    fitz.Point(separator_x, 14),
                    fitz.Point(separator_x, output_height - 14),
                    color=(0.65, 0.68, 0.72),
                    width=0.6,
                )

                if not translations:
                    continue
                left = source_rect.width + 20
                right = source_rect.width + panel_width - 12
                chars_per_line = max(8, int((right - left) / font_size))
                top = 18.0
                for translation in translations:
                    line_count = max(1, math.ceil(len(translation) / chars_per_line))
                    height = line_count * font_size * 1.45 + 8
                    bottom = top + height
                    if bottom > output_height - 15:
                        raise RuntimeError(
                            "译文侧栏高度估算不足，已停止生成，避免输出缺字 PDF。"
                        )
                    spare_height = output_page.insert_textbox(
                        fitz.Rect(left, top, right, bottom),
                        translation,
                        fontname="china-s",
                        fontsize=font_size,
                        lineheight=1.2,
                        color=(0.78, 0.10, 0.12),
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                    if spare_height < 0:
                        raise RuntimeError(
                            "译文侧栏渲染空间不足，已停止生成，避免输出缺字 PDF。"
                        )
                    top = bottom + 3
        if output.page_count == 0:
            raise ValueError("PDF 没有可渲染页面。")
        output.save(output_pdf, garbage=4, deflate=True)
    finally:
        output.close()
