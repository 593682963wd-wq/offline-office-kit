"""作者：Wangyuan；日期：2026-10-29。按调用方表头填写文档，不附带模板。"""

from __future__ import annotations

import re
from collections.abc import Iterable, Sequence
from copy import deepcopy
from io import BytesIO
from typing import BinaryIO

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def _normal_header(value: str) -> str:
    return re.sub(r"\s+", "", value)


def fill_table(
    source: str | BinaryIO,
    *,
    headers: Sequence[str],
    rows: Iterable[Sequence[object]],
    mode: str = "replace",
) -> BytesIO:
    """填写唯一匹配的表格，支持明确替换或追加，不猜测日期及保留制度。

    保留表头及数据行原有布局；数据统一居中并以普通文本写入。
    若存在两个相同表头的表，调用方应先准备无歧义模板。
    """
    if mode not in {"replace", "append"}:
        raise ValueError("mode 只能为 replace 或 append")
    if not headers or any(not isinstance(h, str) or not h.strip() for h in headers):
        raise ValueError("表头必须为非空字符串序列")
    expected = [_normal_header(h) for h in headers]
    if len(set(expected)) != len(expected):
        raise ValueError("表头不能重复")
    values = [list(row) for row in rows]
    if any(len(row) != len(headers) for row in values):
        raise ValueError("每一行列数必须与表头完全一致")
    document = Document(source)
    matches = [
        table
        for table in document.tables
        if table.rows
        and [_normal_header(c.text) for c in table.rows[0].cells] == expected
    ]
    if len(matches) != 1:
        raise ValueError("未找到唯一匹配表格，请核对表头和模板")
    table = matches[0]
    # 先校验结构再复制/删行，避免嵌套表或内容控件中的旧文字随原型复制。
    if any(
        child.tag not in {qn("w:tblPr"), qn("w:tblGrid"), qn("w:tr")}
        for child in table._tbl
    ):
        raise ValueError("目标表包含不支持的内容控件或扩展结构")
    for row in table.rows:
        if any(child.tag not in {qn("w:trPr"), qn("w:tc")} for child in row._tr):
            raise ValueError("目标表行包含不支持的内容控件或扩展结构")
        if len({id(cell._tc) for cell in row.cells}) != len(headers):
            raise ValueError("目标表不能包含合并单元格")
        for cell in row.cells:
            if cell._tc.xpath("./w:tcPr/w:vMerge | ./w:tcPr/w:gridSpan"):
                raise ValueError("目标表不能包含合并单元格")
            if any(child.tag not in {qn("w:tcPr"), qn("w:p")} for child in cell._tc):
                raise ValueError("目标单元格不能包含嵌套表、内容控件或其他非段落结构")
    prototype = deepcopy(table.rows[1]._tr) if len(table.rows) > 1 else None
    if mode == "replace":
        for row in list(table.rows[1:]):
            table._tbl.remove(row._tr)
    for row_values in values:
        if prototype is None:
            row = table.add_row()
        else:
            table._tbl.append(deepcopy(prototype))
            row = table.rows[-1]
        # 合并单元格会造成多个列写入同一底层节点，宁可拒绝而不静默丢值。
        if len({id(cell._tc) for cell in row.cells}) != len(headers):
            raise ValueError("数据行不能包含合并单元格")
        for cell, value in zip(row.cells, row_values):
            paragraphs = list(cell.paragraphs)
            paragraph = paragraphs[0] if paragraphs else cell.add_paragraph()
            for extra in paragraphs[1:]:
                extra._element.getparent().remove(extra._element)
            for child in list(paragraph._element):
                if child.tag != qn("w:pPr"):
                    paragraph._element.remove(child)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            paragraph.add_run("" if value is None else str(value))
    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output
