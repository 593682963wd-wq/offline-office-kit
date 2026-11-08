"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""

from __future__ import annotations

from xml.sax.saxutils import escape

WHITE = "FFFFFF"

TEXT = "111111"

def emu(v: float) -> int:
    """英寸 -> EMU。"""
    return int(v * 914400)

def rr(text: str, sz: int, color: str = TEXT, bold: bool = False, italic: bool = False) -> str:
    """单个 run。sz 单位为 OOXML 的百倍磅（如 1800 = 18pt）。"""
    attrs = f' sz="{sz}"'
    if bold:
        attrs += ' b="1"'
    if italic:
        attrs += ' i="1"'
    return (
        f'<a:r><a:rPr lang="zh-CN" altLang="en-US"{attrs}>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        '<a:latin typeface="Arial"/><a:ea typeface="Arial"/>'
        f'</a:rPr><a:t>{escape(text)}</a:t></a:r>'
    )

def para(runs: list[str], align: str = "ctr", line_spc_pct: int = 112000) -> str:
    """一段。align: l / ctr / r / just。"""
    return (
        f'<a:p><a:pPr algn="{align}"><a:lnSpc><a:spcPct val="{line_spc_pct}"/></a:lnSpc></a:pPr>'
        + "".join(runs)
        + "</a:p>"
    )

class Builder:
    """单张 slide 的 shape 累积器，最后 .xml() 输出整张 slide XML。"""

    def __init__(self) -> None:
        self._next_id = 2
        self.parts: list[str] = []

    def _sid(self) -> int:
        x = self._next_id
        self._next_id += 1
        return x

    def sp(
        self,
        name: str,
        x: int,
        y: int,
        cx: int,
        cy: int,
        fill: str | None,
        line: str | None,
        text: str = "",
        txbox: bool = False,
        geom: str = "rect",
        anchor: str = "ctr",
    ) -> None:
        sid = self._sid()
        fa = "<a:noFill/>" if fill is None else f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        la = "<a:ln><a:noFill/></a:ln>" if line is None else (
            f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
        )
        tb = ' txBox="1"' if txbox else ""
        self.parts.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{escape(name, {chr(34): "&quot;", chr(39): "&apos;"})}"/>'
            f'<p:cNvSpPr{tb}/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>{fa}{la}</p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="{anchor}"/>'
            f'<a:lstStyle/>{text}</p:txBody></p:sp>'
        )

    def pic(self, name: str, x: int, y: int, cx: int, cy: int, rid: str = "rId2") -> None:
        sid = self._sid()
        self.parts.append(
            f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="{escape(name, {chr(34): "&quot;", chr(39): "&apos;"})}"/>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        )

    def xml(self, bg: str = WHITE) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'showMasterSp="0">'
            f'<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>'
            '</p:bgPr></p:bg>'
            '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(self.parts)
            + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" '
            'nodeType="tmRoot"/></p:par></p:tnLst></p:timing></p:sld>'
        )
