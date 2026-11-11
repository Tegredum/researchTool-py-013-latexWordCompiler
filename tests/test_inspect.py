# author: Codex, DSH
from pathlib import Path
import sys
import tempfile
import unittest
from zipfile import ZipFile

# Source checkout: allow running this file directly without installing the package.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from latex2word.inspect_docx import CFB_SIGNATURE, InspectionError, inspect_docx


NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
	'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
	'xmlns:o="urn:schemas-microsoft-com:office:office" '
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


class InspectionTests(unittest.TestCase):
	def setUp(self):
		self.temporary = tempfile.TemporaryDirectory()
		self.root = Path(self.temporary.name)
		self.addCleanup(self.temporary.cleanup)

	def package(self, body="", extra=None):
		path = self.root / "sample.docx"
		with ZipFile(path, "w") as archive:
			archive.writestr("word/document.xml", f"<w:document {NS}><w:body>{body}</w:body></w:document>")
			for name, data in (extra or {}).items():
				archive.writestr(name, data)
		return path

	def test_fragments_fields_headers_and_omml_are_audited(self):
		body = ('<w:p><w:r><w:t>L2W_</w:t></w:r><w:r><w:t>MATH_001 </w:t></w:r>'
				'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
				'<w:r><w:instrText> SE</w:instrText></w:r>'
				'<w:r><w:instrText>Q Equation \\* ARABIC </w:instrText></w:r>'
				'<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
				'<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
				'<m:oMathPara><m:oMath/></m:oMathPara></w:p>')
		header = f'<w:hdr {NS}><w:p><w:fldSimple w:instr="REF L2W_REF_002"/></w:p></w:hdr>'
		report = inspect_docx(self.package(body, {"word/header1.xml": header}))
		self.assertEqual(report["counts"]["paragraphs"], 2)
		self.assertEqual(report["counts"]["omml"], 1)
		self.assertEqual(report["counts"]["fields"], 2)
		self.assertEqual(report["counts"]["independent_equation_sequences"], 1)
		self.assertEqual({item["token"] for item in report["placeholders"]}, {"L2W_MATH_001", "L2W_REF_002"})
		self.assertFalse(report["visual_layout_verified"])

	def test_ole_requires_real_relationship_and_compound_signature(self):
		body = ('<w:p><w:r><w:object><o:OLEObject ProgID="Equation.DSMT4" r:id="r1"/>'
				'<o:OLEObject ProgID="Equation.DSMT4" r:id="missing"/>'
				'<o:OLEObject ProgID="Equation.3" r:id="r2"/>'
				'<o:OLEObject ProgID="Equation.DSMT4" r:id="r3"/></w:object></w:r></w:p>')
		relationships = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
						'<Relationship Id="r1" Target="embeddings/a.bin"/>'
						'<Relationship Id="r2" Target="embeddings/b.bin"/>'
						'<Relationship Id="r3" Target="embeddings/c.bin"/>'
						'</Relationships>')
		report = inspect_docx(self.package(body, {
			"word/_rels/document.xml.rels": relationships,
			"word/embeddings/a.bin": CFB_SIGNATURE + b"fixture",
			"word/embeddings/b.bin": CFB_SIGNATURE + b"fixture",
			"word/embeddings/c.bin": b"not an OLE compound file",
		}))
		self.assertEqual(report["counts"]["ole_objects"], 4)
		self.assertEqual(report["counts"]["mathtype_objects"], 1)
		self.assertEqual(report["counts"]["embeddings"], 3)
		self.assertFalse(report["ole_objects"][1]["embedded_exists"])
		self.assertFalse(report["ole_objects"][2]["declared_mathtype"])
		self.assertFalse(report["ole_objects"][3]["compound_file"])
		self.assertFalse(report["mathtype_editability_verified"])

	def test_styles_inheritance_direct_formatting_sections_and_bookmarks(self):
		styles = (f'<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
				'<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>'
				'<w:style w:type="paragraph" w:styleId="Equation"><w:name w:val="公式"/><w:basedOn w:val="Normal"/>'
				'<w:pPr><w:jc w:val="left"/><w:ind w:firstLine="0"/></w:pPr></w:style></w:styles>')
		body = ('<w:p><w:pPr><w:pStyle w:val="Equation"/><w:ind w:firstLine="480"/></w:pPr>'
				'<w:bookmarkStart w:id="1" w:name="eq_a"/><w:r><w:rPr><w:b/></w:rPr><w:t>$x$</w:t></w:r>'
				'<w:bookmarkEnd w:id="1"/></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>')
		report = inspect_docx(self.package(body, {"word/styles.xml": styles}))
		style = next(item for item in report["styles"] if item["style_id"] == "Equation")
		self.assertEqual(style["inheritance_chain"], ["Normal"])
		self.assertEqual(style["paragraph_properties"]["ind"]["firstLine"], "0")
		self.assertEqual(style["samples"][0]["direct_paragraph_properties"]["ind"]["firstLine"], "480")
		self.assertEqual(style["samples"][0]["direct_run_properties"], [{"b": {}}])
		self.assertEqual(report["bookmarks"][0]["name"], "eq_a")
		self.assertEqual(report["sections"][0]["properties"]["pgSz"]["w"], "11906")
		self.assertEqual(report["counts"]["tex_delimited_spans"], 1)

	def test_mathtype_native_sequence_is_not_independent_equation_sequence(self):
		body = '<w:p><w:fldSimple w:instr="SEQ MTEqn"/><w:fldSimple w:instr="SEQ EquationOther"/></w:p>'
		report = inspect_docx(self.package(body))
		self.assertEqual(report["counts"]["independent_equation_sequences"], 0)

	def test_corrupt_packages_and_xml_fail_clearly(self):
		path = self.root / "bad.docx"
		path.write_text("not a package", encoding="utf-8")
		with self.assertRaisesRegex(InspectionError, "Cannot open"):
			inspect_docx(path)
		with self.assertRaisesRegex(InspectionError, "Cannot parse"):
			inspect_docx(self.package("<broken>"))

	def test_inheritance_cycles_and_unclosed_fields_are_reported(self):
		styles = (f'<w:styles {NS}><w:style w:styleId="A"><w:basedOn w:val="B"/></w:style>'
				'<w:style w:styleId="B"><w:basedOn w:val="A"/></w:style></w:styles>')
		body = '<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>REF missing</w:instrText></w:r></w:p>'
		report = inspect_docx(self.package(body, {"word/styles.xml": styles}))
		self.assertTrue(any("Cyclic" in warning for warning in report["warnings"]))
		self.assertTrue(any("Unclosed" in warning for warning in report["warnings"]))


if __name__ == "__main__":
	unittest.main()
