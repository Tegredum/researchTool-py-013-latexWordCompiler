# encoding: utf-8
# author: Codex, DSH
import json
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET

# Source checkout: allow running this file directly without installing the package.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from latex2word.config import load_config
from latex2word.writer import NS, W, WriterError, style_docx


def document(path, text, extra=""):
	parts = {
		"word/document.xml": f'<w:document xmlns:w="{W}"><w:body>{text}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/>{extra}</w:sectPr></w:body></w:document>',
		"word/styles.xml": f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:color w:val="123456"/></w:rPr></w:style></w:styles>',
		"word/numbering.xml": f'<w:numbering xmlns:w="{W}"/>',
		"word/settings.xml": f'<w:settings xmlns:w="{W}"/>',
	}
	with ZipFile(path, "w") as z:
		for name, data in parts.items():
			z.writestr(name, data.encode())


def paragraph(text):
	return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'


def xml(path, part="word/document.xml"):
	with ZipFile(path) as z:
		return ET.fromstring(z.read(part))


class WriterTests(unittest.TestCase):
	def setUp(self):
		self.temp = tempfile.TemporaryDirectory()
		self.addCleanup(self.temp.cleanup)
		self.base = Path(self.temp.name)
		self.input, self.output = self.base / "in.docx", self.base / "out.docx"

	def test_inline_tex_exact_bookmark_and_surrounding_text(self):
		document(self.input, paragraph("中L2W_MATH_eq0001文 English words"))
		manifest = {"equations": [{"id": "eq0001", "token": "L2W_MATH_eq0001", "bookmark": "L2WEq0001", "tex": r"\mathbf{x}+x_i", "display": False, "numbered": False}]}
		style_docx(self.input, self.output, manifest, load_config())
		root = xml(self.output)
		self.assertEqual("".join(root.xpath(".//w:t/text()", namespaces=NS)), "中$\\mathbf{x}+x_i$文 English words")
		start = root.find(".//w:bookmarkStart", NS)
		self.assertEqual(start.getnext().find("w:t", NS).text, r"$\mathbf{x}+x_i$")
		self.assertEqual(start.getnext().getnext().tag, f"{{{W}}}bookmarkEnd")
		self.assertFalse(root.xpath(".//*[local-name()='oMath']"))

	def test_display_uses_two_tabs_and_separate_native_number_marker(self):
		document(self.input, paragraph("L2W_MATH_eq0001"))
		eq = {"id": "eq0001", "token": "L2W_MATH_eq0001", "bookmark": "L2WEq0001", "tex": "x=1", "display": True, "numbered": True}
		style_docx(self.input, self.output, {"equations": [eq]}, load_config())
		root = xml(self.output)
		self.assertEqual(len(root.findall(".//w:r/w:tab", NS)), 2)
		self.assertIn("L2W_NUMBER_eq0001", root.xpath(".//w:t/text()", namespaces=NS))
		self.assertFalse(root.findall(".//w:instrText", NS))
		para = root.find(".//w:p", NS)
		self.assertEqual(para.find("w:pPr/w:jc", NS).get(f"{{{W}}}val"), "left")

	def test_captions_use_fields_and_refs_to_unique_bookmarks(self):
		document(self.input, paragraph("L2W_CAP_fig0001 A figure") + paragraph("See L2W_REF_0001"))
		manifest = {"targets": [{"kind": "figure", "label": "fig:a", "bookmark": "L2WFig0001", "token": "L2W_CAP_fig0001"}], "references": [{"kind": "figure", "target": "fig:a", "bookmark": "L2WRef0001", "token": "L2W_REF_0001"}]}
		style_docx(self.input, self.output, manifest, load_config())
		codes = xml(self.output).xpath(".//w:instrText/text()", namespaces=NS)
		self.assertIn(" SEQ Figure \\* ARABIC ", codes)
		self.assertIn(" REF L2WFig0001 \\h ", codes)
		self.assertFalse(any("Equation" in c for c in codes))

	def test_reference_style_inheritance_not_overwritten(self):
		document(self.input, paragraph("text"))
		config_file = self.base / "config.json"
		config_file.write_text(json.dumps({"version": 1, "reference_docx": "in.docx"}), encoding="utf-8")
		style_docx(self.input, self.output, {}, load_config(config_file))
		normal = xml(self.output, "word/styles.xml").find("w:style[@w:styleId='Normal']", NS)
		self.assertEqual(normal.find("w:rPr/w:color", NS).get(f"{{{W}}}val"), "123456")
		self.assertIsNone(normal.find("w:rPr/w:sz", NS))

	def test_display_in_columns_fails_explicitly(self):
		document(self.input, paragraph("L2W_MATH_eq0001"), '<w:cols w:num="2"/>')
		eq = {"id": "eq0001", "token": "L2W_MATH_eq0001", "bookmark": "L2WEq0001", "tex": "x", "display": True, "numbered": False}
		with self.assertRaisesRegex(WriterError, "single-column"):
			style_docx(self.input, self.output, {"equations": [eq]}, load_config())
		self.assertFalse(self.output.exists())

	def test_duplicate_tokens_do_not_silently_replace_one(self):
		document(self.input, paragraph("L2W_MATH_eq0001") + paragraph("L2W_MATH_eq0001"))
		eq = {"id": "eq0001", "token": "L2W_MATH_eq0001", "bookmark": "L2WEq0001", "tex": "x", "display": False, "numbered": False}
		with self.assertRaises(WriterError):
			style_docx(self.input, self.output, {"equations": [eq]}, load_config())

	def test_title_metadata_does_not_leak_internal_math_token(self):
		document(self.input, paragraph("L2W_MATH_eq0001"))
		with ZipFile(self.input, "a") as package:
			package.writestr("docProps/core.xml", '<properties><title>Model L2W_MATH_eq0001</title></properties>')
		eq = {"id": "eq0001", "token": "L2W_MATH_eq0001", "bookmark": "L2WEq0001", "tex": "x", "display": False, "numbered": False}
		style_docx(self.input, self.output, {"equations": [eq]}, load_config())
		self.assertEqual(xml(self.output, "docProps/core.xml").find("title").text, "Model $x$")


if __name__ == "__main__":
	unittest.main()
