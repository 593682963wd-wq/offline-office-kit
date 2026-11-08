"""作者：Wangyuan；日期：2026-10-29。合成可编辑文件与抽取回读。"""
import unittest,tempfile,io,json
from pathlib import Path
from datetime import date
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.oxml.xmlchemy import OxmlElement
import fitz
from office import ppt_edit,pdf_overlay,pdf_text,pdf_metadata,date_counts,scenario_dates,scenario_overlay,section_diff
class OfficeExtensionTests(unittest.TestCase):
    def test_ppt_empty_paragraph(self):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);tf=slide.shapes.add_textbox(Inches(1),Inches(1),Inches(3),Inches(1)).text_frame
        ppt_edit.set_multi_paragraphs(tf,lambda text:False,["合成替换"]);self.assertIn("合成替换",tf.text)
    def test_ppt_readback_format(self):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);shape=slide.shapes.add_textbox(Inches(1),Inches(1),Inches(3),Inches(1));shape.text="old";shape.text_frame.paragraphs[0].runs[0].font.size=Pt(18)
        ppt_edit.set_text_keep_format(shape.text_frame,"合成文字");stream=io.BytesIO();prs.save(stream);loaded=Presentation(io.BytesIO(stream.getvalue()));self.assertEqual(loaded.slides[0].shapes[0].text,"合成文字");self.assertEqual(loaded.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.size,Pt(18))
    def test_ppt_hidden_field_cleared(self):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);tf=slide.shapes.add_textbox(Inches(1),Inches(1),Inches(3),Inches(1)).text_frame
        fld=OxmlElement("a:fld");node=OxmlElement("a:t");node.text="old hidden field";fld.append(node);tf.paragraphs[0]._p.append(fld)
        ppt_edit.set_text_keep_format(tf,"new");self.assertNotIn("old hidden field",tf._txBody.xml)
    def test_ppt_delete(self):
        prs=Presentation();prs.slides.add_slide(prs.slide_layouts[6]);prs.slides.add_slide(prs.slide_layouts[6]);ppt_edit.delete_slides_after(prs,1);self.assertEqual(len(prs.slides),1)
        with self.assertRaises(ValueError):ppt_edit.delete_slides_after(prs,-1)
    def test_ppt_clone(self):
        prs=Presentation();slide=prs.slides.add_slide(prs.slide_layouts[6]);slide.shapes.add_textbox(Inches(1),Inches(1),Inches(3),Inches(1)).text="synthetic";ppt_edit.clone_slide_after(prs,0);stream=io.BytesIO();prs.save(stream);self.assertEqual(len(Presentation(io.BytesIO(stream.getvalue())).slides),2)
    def test_pdf_text(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/"synthetic.pdf";doc=fitz.open();doc.new_page().insert_text((50,50),"Synthetic evidence 123");doc.save(path);doc.close()
            text,warnings=pdf_text._pdf_to_markdown_with_warnings(path);self.assertIn("Synthetic evidence",text)
    def test_dates(self):
        self.assertEqual(scenario_dates._exec_in_window("1234567",date(2040,1,1),date(2040,1,7)),7)
    def test_diff(self):self.assertNotEqual(section_diff.compare_markdown("## A\nold","## A\nnew"),[])
    def test_pdf_overlay_readback(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);source=root/"source.pdf";template=root/"translation.json";output=root/"result.pdf"
            doc=fitz.open();doc.new_page().insert_text((50,50),"Synthetic source text");doc.save(source);doc.close()
            payload=pdf_overlay.generate_translation_template(source,template)
            data=json.loads(template.read_text());self.assertTrue(data["entries"]);data["entries"][0]["translation"]="Synthetic translated text";template.write_text(json.dumps(data))
            pdf_overlay.render_bilingual_pdf(source,output,template)
            with fitz.open(output) as doc:self.assertEqual(len(doc),1);self.assertIn("Synthetic translated",doc[0].get_text());self.assertIn("Synthetic source",doc[0].get_text())
    def test_diff_preserves_decimal(self):
        self.assertTrue(section_diff.compare_markdown("Section 1\n10.5","Section 1\n105"));self.assertFalse(section_diff.compare_markdown("Section 1\nA  B","Section 1\nA B"))

if __name__=="__main__":unittest.main()
