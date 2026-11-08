"""作者：Wangyuan；日期：2026-10-29。离线通用核心，输入由调用方显式提供。"""
from __future__ import annotations
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
import copy

def _p_text(p_elem) -> str:
    return ''.join((t.text or '' for t in p_elem.iter(qn('a:t'))))

def _set_p_text(p_elem, text):
    first = p_elem.find(qn('a:r'))
    run = p_elem.makeelement(qn('a:r'), {})
    if first is not None:
        props = first.find(qn('a:rPr'))
        if props is not None:
            props = copy.deepcopy(props)
            for child in list(props):
                if child.tag in {qn('a:hlinkClick'), qn('a:hlinkMouseOver')}:
                    props.remove(child)
            run.append(props)
    for child in list(p_elem):
        if child.tag not in {qn('a:pPr'), qn('a:endParaRPr')}:
            p_elem.remove(child)
    node = p_elem.makeelement(qn('a:t'), {})
    node.text = str(text)
    run.append(node)
    end = p_elem.find(qn('a:endParaRPr'))
    if end is not None:
        p_elem.insert(list(p_elem).index(end), run)
    else:
        p_elem.append(run)

def set_text_keep_format(tf, text):
    _set_p_text(tf.paragraphs[0]._p, text)
    body = tf._txBody
    for paragraph in body.findall(qn('a:p'))[1:]:
        body.remove(paragraph)

def set_multi_paragraphs(tf, keep_pred, lines):
    """保留满足 keep_pred 的段（标题），其余内容段用 lines 重建（格式克隆自首个内容段）。"""
    body = tf._txBody
    ps = body.findall(qn('a:p'))
    content = [p for p in ps if not keep_pred(_p_text(p))]
    template = copy.deepcopy(content[0]) if content else copy.deepcopy(ps[-1])
    for p in content:
        body.remove(p)
    for line in lines:
        newp = copy.deepcopy(template)
        _set_p_text(newp, line)
        body.append(newp)
    if not body.findall(qn('a:p')):
        _set_p_text(template, '')
        body.append(template)

def set_cell(cell, text):
    set_text_keep_format(cell.text_frame, str(text))

def remove_shape(shape):
    shape._element.getparent().remove(shape._element)

def delete_slides_after(prs, keep):
    if type(keep) is not int or keep < 0:
        raise ValueError('保留页数必须非负整数')
    '删除第 keep 张之后的所有幻灯片（保留前 keep 张）。'
    lst = prs.slides._sldIdLst
    for sld in list(lst)[keep:]:
        prs.part.drop_rel(sld.rId)
        lst.remove(sld)

def clone_slide_after(prs, source_index):
    """Clone a template slide in-place, including its image relationships."""
    source = list(prs.slides)[source_index]
    cloned = prs.slides.add_slide(source.slide_layout)
    for shape in list(cloned.shapes):
        remove_shape(shape)
    rel_ids = {}
    for relationship in source.part.rels.values():
        if relationship.reltype == RT.SLIDE_LAYOUT:
            continue
        target = relationship.target_ref if relationship.is_external else relationship.target_part
        rel_ids[relationship.rId] = cloned.part.relate_to(target, relationship.reltype, is_external=relationship.is_external)
    for shape in source.shapes:
        element = copy.deepcopy(shape.element)
        for child in element.iter():
            for attr in (qn('r:embed'), qn('r:link')):
                old_id = child.get(attr)
                if old_id in rel_ids:
                    child.set(attr, rel_ids[old_id])
        cloned.shapes._spTree.insert_element_before(element, 'p:extLst')
    sld_ids = prs.slides._sldIdLst
    inserted_id = sld_ids[-1]
    sld_ids.remove(inserted_id)
    sld_ids.insert(source_index + 1, inserted_id)
    return list(prs.slides)[source_index + 1]
