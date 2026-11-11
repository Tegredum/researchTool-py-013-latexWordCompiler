# encoding: utf-8
# author: Codex, DSH
import json
from pathlib import Path
import sys
import tempfile
import unittest

# Source checkout: allow running this file directly without installing the package.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from latex2word.config import ConfigError, DEFAULT_CONFIG, load_config

class ConfigTests(unittest.TestCase):
	def setUp(self):
		self.temporary = tempfile.TemporaryDirectory()
		self.root = Path(self.temporary.name)
		self.addCleanup(self.temporary.cleanup)

	def load(self, value):
		path = self.root / "style.json"
		path.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
		return load_config(path)

	def test_defaults_are_independent_and_have_three_heading_levels(self):
		first = load_config()
		first["styles"]["body"]["size_pt"] = 99
		self.assertEqual(load_config()["styles"]["body"]["size_pt"], 12)
		self.assertEqual(DEFAULT_CONFIG["styles"]["body"]["size_pt"], 12)
		self.assertNotIn("heading4", first["styles"])

	def test_reference_is_relative_to_config_and_overrides_remain_explicit(self):
		(self.root / "reference.docx").write_bytes(b"reference used only for path resolution")
		result = self.load({"reference_docx": "reference.docx", "page": {"margin_top_cm": 3},
							"styles": {"body": {"name": "正文", "size_pt": 11}}})
		self.assertEqual(result["reference_docx"], str((self.root / "reference.docx").resolve()))
		self.assertEqual(result["_explicit"]["page"], ["margin_top_cm"])
		self.assertEqual(result["_explicit"]["styles"], {"body": ["name", "size_pt"]})
		self.assertEqual(result["styles"]["body"]["latin_font"], "Times New Roman")
		self.assertNotIn("heading1", result["_explicit"]["styles"])

	def test_invalid_values_fail_with_location(self):
		cases = [
			({"verison": 1}, "unknown key"),
			({"version": True}, "version"),
			({"version": 1.0}, "version"),
			({"version": 2}, "version"),
			({"page": []}, "page"),
			({"page": {"width_cm": True}}, "page.width_cm"),
			({"page": {"margin_left_cm": 20}}, "usable width"),
			({"page": {"height_cm": float("nan")}}, "page.height_cm"),
			({"page": {"height_cm": 10 ** 1000}}, "page.height_cm"),
			({"page": {"margin_top_cm": -1}}, "page.margin_top_cm"),
			({"styles": {"heading4": {}}}, "heading4"),
			({"styles": {"body": {"font": "Arial"}}}, "font"),
			({"styles": {"body": {"bold": 1}}}, "styles.body.bold"),
			({"styles": {"body": {"size_pt": "12"}}}, "styles.body.size_pt"),
			({"styles": {"body": {"name": " "}}}, "styles.body.name"),
			({"styles": {"body": {"latin_font": "A\nB"}}}, "latin_font"),
			({"styles": {"body": {"alignment": []}}}, "alignment"),
			({"styles": {"body": {"line_spacing": 0}}}, "line_spacing"),
			({"headings": {"numbered": "false"}}, "headings.numbered"),
			({"captions": {"figure_label": ""}}, "captions.figure_label"),
			({"math": {"timeout_seconds": 0}}, "math.timeout_seconds"),
			({"math": {"timeout_seconds": 180.5}}, "math.timeout_seconds"),
			({"reference_docx": "missing.docx"}, "does not exist"),
			({"reference_docx": False}, "reference_docx"),
			({"_explicit": {}}, "unknown key"),
		]
		for value, message in cases:
			with self.subTest(value=value):
				with self.assertRaisesRegex(ConfigError, message):
					self.load(value)

	def test_duplicate_json_keys_and_invalid_root_are_rejected(self):
		path = self.root / "bad.json"
		path.write_text('{"styles":{"body":{"size_pt":12,"size_pt":14}}}', encoding="utf-8")
		with self.assertRaisesRegex(ConfigError, "duplicate JSON key: size_pt"):
			load_config(path)
		with self.assertRaisesRegex(ConfigError, "must be a JSON object"):
			self.load([])

	def test_shipped_configuration_matches_defaults(self):
		path = Path(__file__).resolve().parents[1] / "templates" / "default.json"
		result = load_config(path)
		result.pop("_explicit")
		self.assertEqual(result, DEFAULT_CONFIG)

	def test_utf8_bom_is_accepted(self):
		path = self.root / "bom.json"
		path.write_text('{"captions":{"figure_label":"图"}}', encoding="utf-8-sig")
		self.assertEqual(load_config(path)["captions"]["figure_label"], "图")


if __name__ == "__main__":
	unittest.main()
