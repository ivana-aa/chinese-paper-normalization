"""Regression checks for the read-only skill helper, using temporary DOCX fixtures."""

import hashlib
import importlib.util
import json
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile


SKILL = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("audit_docx", SKILL / "scripts" / "audit_docx.py")
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)
PROFILE = json.loads((SKILL / "references" / "default-profile.json").read_text(encoding="utf-8"))

STYLES = f'''<w:styles xmlns:w="{MODULE.NS['w']}">
<w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:eastAsia="宋体" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
<w:sz w:val="24"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:jc w:val="both"/>
<w:spacing w:line="360" w:lineRule="auto" w:before="0" w:after="0"/>
<w:ind w:firstLine="480"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/>
<w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="left"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="21"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>
</w:styles>'''


def paragraph(text, ppr="", rpr=""):
    return f"<w:p><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t>{text}</w:t></w:r></w:p>"


TITLE = paragraph(
    "1研究内容", '<w:jc w:val="center"/>',
    '<w:rFonts w:eastAsia="黑体"/><w:sz w:val="36"/><w:b/>',
)
BODY = paragraph("采用U-Net完成分割。")
CELL = f'''<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr>
{paragraph('Dice', '<w:jc w:val="center"/>', '<w:sz w:val="21"/>')}</w:tc>'''
TABLE = f'''<w:tbl><w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="12"/><w:bottom w:val="single" w:sz="12"/>
<w:left w:val="nil"/><w:right w:val="nil"/><w:insideV w:val="nil"/>
</w:tblBorders></w:tblPr><w:tr>{CELL}</w:tr></w:tbl>'''
SECTION = '''<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1440" w:bottom="1440" w:left="1797" w:right="1797"/></w:sectPr>'''


class AuditTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent)
        self.assertTrue(Path(self.temporary.name).resolve().is_relative_to(Path(__file__).resolve().parent))
        self.addCleanup(self.temporary.cleanup)

    def run_audit(self, content, profile=None):
        path = Path(self.temporary.name) / "fixture.docx"
        document = f'''<w:document xmlns:w="{MODULE.NS['w']}" xmlns:m="{MODULE.NS['m']}">
<w:body>{content}{SECTION}</w:body></w:document>'''
        with ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", document)
            archive.writestr("word/styles.xml", STYLES)
            archive.writestr("word/media/fixture.bin", b"unaltered-media-fixture")
        before = path.read_bytes()
        report = MODULE.audit(path, profile or PROFILE)
        self.assertEqual(before, path.read_bytes(), "The audit must never rewrite its input")
        self.assertEqual(hashlib.sha256(before).hexdigest(), report["sha256"])
        return report

    def test_inherited_fonts_and_direct_formats(self):
        report = self.run_audit(TITLE + BODY + TABLE)
        self.assertEqual([], report["issues"])
        self.assertEqual(1, len(report["tables"]))
        self.assertTrue(report["tables"][0]["requires_visual_review"])

    def test_wrong_manual_heading_font_is_found(self):
        heading = paragraph("3.1模型流程", '<w:jc w:val="left"/>', '<w:sz w:val="32"/><w:b/>')
        report = self.run_audit(TITLE + heading)
        errors = [i for i in report["issues"] if i["check"] == "chinese_font"]
        self.assertEqual(1, len(errors))
        self.assertEqual("heading2", errors[0]["role"])
        self.assertEqual("黑体", errors[0]["expected"])

    def test_boundary_spaces_across_runs_and_code_exception(self):
        mixed = '<w:p><w:r><w:t xml:space="preserve">采用 </w:t></w:r>'
        mixed += '<w:r><w:t>U-Net模型</w:t></w:r></w:p>'
        english = paragraph("Attention U-Net")
        code = paragraph("if x &gt; 0:  # 中文 comment", '<w:pStyle w:val="Code"/>')
        report = self.run_audit(TITLE + mixed + english + code)
        spaces = [i for i in report["issues"] if i["check"] == "cjk_latin_space_candidate"]
        self.assertEqual(1, len(spaces))
        self.assertEqual("采用 U-Net模型", spaces[0]["text"])
        self.assertFalse(any(i.get("role") == "code" for i in report["issues"]))

    def test_native_math_fields_and_media_inventory(self):
        equation = '<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>'
        field = '<w:p><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r></w:p>'
        report = self.run_audit(TITLE + equation + field)
        self.assertEqual(1, report["native_math_count"])
        self.assertEqual([" PAGE "], report["field_instructions"])
        self.assertEqual(
            hashlib.sha256(b"unaltered-media-fixture").hexdigest(),
            report["media_sha256"]["word/media/fixture.bin"],
        )

    def test_visible_vertical_border_is_only_a_candidate(self):
        grid = TABLE.replace('<w:insideV w:val="nil"/>', '<w:insideV w:val="single" w:sz="4"/>')
        report = self.run_audit(TITLE + grid)
        self.assertEqual(1, report["checks"]["visible_vertical_border_candidate"])

    def test_numbers_and_toc_do_not_become_headings(self):
        toc = paragraph("1研究内容", '<w:pStyle w:val="TOC1"/>')
        report = self.run_audit(toc + TITLE + paragraph("512×512影像") + paragraph("0.94581为均值"))
        self.assertEqual(["1研究内容"], [h["text"] for h in report["headings"]])

    def test_profile_override_and_character_indent_precedence(self):
        heading = paragraph("3.1模型流程", '<w:jc w:val="left"/>', '<w:sz w:val="24"/><w:b/>')
        custom = MODULE.merge_profile(PROFILE, {"heading2": {"chinese_font": "宋体", "size_pt": 12}})
        self.assertEqual([], self.run_audit(TITLE + heading, custom)["issues"])
        wrong_indent = paragraph("正文缩进测试", '<w:ind w:firstLine="480" w:firstLineChars="100"/>')
        report = self.run_audit(TITLE + wrong_indent)
        self.assertEqual(1, report["checks"]["first_line_indent"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
