# author: Codex
"""Read-only structural audits of DOCX packages, without Microsoft Word.

The report is deliberately not a claim of visual correctness or MathType
round-trip compatibility. Embedded payload signatures and declared ProgIDs are
checked; only Word/MathType can establish that each object opens and edits.
"""

from __future__ import annotations

from collections import Counter
from pathlib import Path, PurePosixPath
import posixpath
import re
from typing import Any
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile


NS = {
	"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	"m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
	"o": "urn:schemas-microsoft-com:office:office",
	"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
CFB_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
PLACEHOLDER_RE = re.compile(r"L2W_[A-Za-z0-9_:\-]+")
EQUATION_SEQUENCE_RE = re.compile(r'(?:^|\s)SEQ\s+(?:"Equation"|Equation)(?=\s|\\|$)', re.I)
MAX_XML_BYTES = 64 * 1024 * 1024


class InspectionError(ValueError):
	"""An input is not a readable Word OOXML document."""


def _q(prefix: str, name: str) -> str:
	return f"{{{NS[prefix]}}}{name}"


def _local(name: str) -> str:
	return name.rsplit("}", 1)[-1]


def _properties(element: ET.Element | None) -> dict[str, Any]:
	"""Retain OOXML values and units, including direct formatting details."""
	if element is None:
		return {}
	result: dict[str, Any] = {_local(k): v for k, v in element.attrib.items()}
	for child in element:
		name = _local(child.tag)
		value = _properties(child)
		if name in result:
			if not isinstance(result[name], list):
				result[name] = [result[name]]
			result[name].append(value)
		else:
			result[name] = value
	return result


def _read_xml(archive: ZipFile, part: str) -> ET.Element:
	if archive.getinfo(part).file_size > MAX_XML_BYTES:
		raise InspectionError(f"XML part exceeds 64 MiB inspection limit: {part}")
	try:
		return ET.fromstring(archive.read(part))
	except (ET.ParseError, RuntimeError, OSError) as exc:
		raise InspectionError(f"Cannot parse {part}: {exc}") from exc


def _paragraph_text(paragraph: ET.Element) -> str:
	# Each paragraph is audited separately, including nested textbox paragraphs.
	# Do not include their text a second time in the outer paragraph's sample.
	chunks: list[str] = []
	def walk(element: ET.Element) -> None:
		for child in element:
			if child.tag == _q("w", "p"):
				continue
			if child.tag == _q("w", "t"):
				chunks.append(child.text or "")
			elif child.tag == _q("w", "tab"):
				chunks.append("\t")
			elif child.tag in {_q("w", "br"), _q("w", "cr")}:
				chunks.append("\n")
			else:
				walk(child)
	walk(paragraph)
	return "".join(chunks)


def _fields(root: ET.Element, part: str, warnings: list[str]) -> list[dict[str, Any]]:
	fields: list[dict[str, Any]] = []
	stack: list[dict[str, Any]] = []
	for element in root.iter():
		if element.tag == _q("w", "fldSimple"):
			fields.append({"part": part, "kind": "simple", "code": element.get(_q("w", "instr"), "").strip()})
		elif element.tag == _q("w", "fldChar"):
			kind = element.get(_q("w", "fldCharType"))
			if kind == "begin":
				stack.append({"chunks": [], "in_code": True})
			elif kind == "separate" and stack:
				stack[-1]["in_code"] = False
			elif kind == "end":
				if stack:
					field = stack.pop()
					fields.append({"part": part, "kind": "complex", "code": "".join(field["chunks"]).strip()})
				else:
					warnings.append(f"Unmatched field end in {part}")
		elif element.tag == _q("w", "instrText") and stack and stack[-1]["in_code"]:
			stack[-1]["chunks"].append(element.text or "")
	for field in stack:
		fields.append({"part": part, "kind": "complex_unclosed", "code": "".join(field["chunks"]).strip()})
		warnings.append(f"Unclosed complex field in {part}")
	return fields


def _relationship_map(archive: ZipFile, names: set[str], part: str) -> dict[str, dict[str, str]]:
	source = PurePosixPath(part)
	rels_name = str(source.parent / "_rels" / (source.name + ".rels"))
	if rels_name not in names:
		return {}
	return {rel.get("Id", ""): dict(rel.attrib) for rel in _read_xml(archive, rels_name)}


def _payload(archive: ZipFile, names: set[str], target: str | None) -> dict[str, Any]:
	exists = bool(target and target in names)
	signature = b""
	if exists:
		with archive.open(target) as stream:
			signature = stream.read(8)
	return {"embedded_exists": exists, "compound_file": signature == CFB_SIGNATURE,
			"payload_kind": "compound" if signature == CFB_SIGNATURE else "zip" if signature.startswith(b"PK\x03\x04") else "unknown"}


def _styles(root: ET.Element | None, samples: dict[str, list[dict[str, Any]]], warnings: list[str]) -> list[dict[str, Any]]:
	if root is None:
		return []
	result: list[dict[str, Any]] = []
	for style in root.findall("w:style", NS):
		style_id = style.get(_q("w", "styleId"), "")
		name = style.find("w:name", NS)
		based_on = style.find("w:basedOn", NS)
		item = {
			"style_id": style_id,
			"name": name.get(_q("w", "val"), "") if name is not None else "",
			"type": style.get(_q("w", "type"), ""),
			"default": style.get(_q("w", "default")) in {"1", "true", "on"},
			"based_on": based_on.get(_q("w", "val")) if based_on is not None else None,
			"paragraph_properties": _properties(style.find("w:pPr", NS)),
			"run_properties": _properties(style.find("w:rPr", NS)),
			"table_properties": _properties(style.find("w:tblPr", NS)),
			"samples": samples.get(style_id, []),
		}
		result.append(item)
	by_id = {item["style_id"]: item for item in result}
	for item in result:
		chain: list[str] = []
		current = item["based_on"]
		seen = {item["style_id"]}
		while current:
			if current in seen:
				warnings.append(f"Cyclic style inheritance at {item['style_id']}: {current}")
				break
			seen.add(current)
			chain.append(current)
			if current not in by_id:
				warnings.append(f"Style {item['style_id']} inherits missing style {current}")
				break
			current = by_id[current]["based_on"]
		item["inheritance_chain"] = chain
	return result


def inspect_docx(path: Path) -> dict[str, Any]:
	"""Audit all Word content parts and return JSON-serializable evidence.

	Paragraph samples are capped at three per style, 160 characters per sample,
	and eight distinct direct run formats. Attributes retain their OOXML units
	(twips, half-points, etc.). ``visual_layout_verified`` is always false.
	"""
	path = Path(path)
	try:
		archive = ZipFile(path)
	except (OSError, BadZipFile) as exc:
		raise InspectionError(f"Cannot open Word package {path}: {exc}") from exc
	with archive:
		all_names = archive.namelist()
		names = set(all_names)
		if "word/document.xml" not in names:
			raise InspectionError("Not a Word OOXML package: word/document.xml is absent")
		warnings: list[str] = []
		if len(names) != len(all_names):
			warnings.append("Package contains duplicate ZIP entry names")
		parts = sorted(name for name in names if name.startswith("word/") and name.endswith(".xml"))
		roots = {part: _read_xml(archive, part) for part in parts}
		styles_root = roots.get("word/styles.xml")
		default_style = "Normal"
		if styles_root is not None:
			for style in styles_root.findall("w:style", NS):
				if style.get(_q("w", "type")) == "paragraph" and style.get(_q("w", "default")) in {"1", "true", "on"}:
					default_style = style.get(_q("w", "styleId"), "Normal")
					break
		fields: list[dict[str, Any]] = []
		bookmarks: list[dict[str, Any]] = []
		ole_objects: list[dict[str, Any]] = []
		placeholders: list[dict[str, Any]] = []
		sections: list[dict[str, Any]] = []
		samples: dict[str, list[dict[str, Any]]] = {}
		style_counts: Counter[str] = Counter()
		counts: Counter[str] = Counter()
		counts.update(dict.fromkeys(("paragraphs", "tables", "omml", "ole_objects", "mathtype_objects",
									"embeddings", "fields", "bookmarks", "placeholders",
									"independent_equation_sequences", "tex_delimited_spans"), 0))
		for part, root in roots.items():
			for node in root.iter():
				for attribute, value in node.attrib.items():
					for match in PLACEHOLDER_RE.finditer(value):
						placeholders.append({"part": part, "attribute": attribute, "token": match.group(0)})
			paragraphs = list(root.iter(_q("w", "p")))
			counts["paragraphs"] += len(paragraphs)
			counts["tables"] += sum(1 for _ in root.iter(_q("w", "tbl")))
			# oMathPara wraps oMath and is not a second formula.
			counts["omml"] += sum(1 for _ in root.iter(_q("m", "oMath")))
			fields.extend(_fields(root, part, warnings))
			for index, paragraph in enumerate(paragraphs):
				ppr = paragraph.find("w:pPr", NS)
				pstyle = paragraph.find("w:pPr/w:pStyle", NS)
				style_id = pstyle.get(_q("w", "val"), default_style) if pstyle is not None else default_style
				style_counts[style_id] += 1
				content = _paragraph_text(paragraph)
				for match in PLACEHOLDER_RE.finditer(content):
					placeholders.append({"part": part, "paragraph_index": index, "token": match.group(0)})
				counts["tex_delimited_spans"] += len(re.findall(r"(?<!\\)\$(?!\$)(?:\\.|[^$\r\n])+?(?<!\\)\$(?!\$)", content))
				bucket = samples.setdefault(style_id, [])
				if len(bucket) < 3:
					run_formats: list[dict[str, Any]] = []
					for run in paragraph.findall(".//w:rPr", NS):
						props = _properties(run)
						if props and props not in run_formats:
							run_formats.append(props)
						if len(run_formats) >= 8:
							break
					bucket.append({"part": part, "paragraph_index": index,
								"text_sample": content[:160], "text_truncated": len(content) > 160,
								"direct_paragraph_properties": _properties(ppr),
								"direct_run_properties": run_formats})
			for bookmark in root.iter(_q("w", "bookmarkStart")):
				bookmarks.append({"part": part, "id": bookmark.get(_q("w", "id")),
								"name": bookmark.get(_q("w", "name"))})
			for section in root.iter(_q("w", "sectPr")):
				sections.append({"part": part, "properties": _properties(section)})
			relationships = _relationship_map(archive, names, part)
			for obj in root.iter(_q("o", "OLEObject")):
				rel_id = obj.get(_q("r", "id"), "")
				rel = relationships.get(rel_id, {})
				target = None
				external = rel.get("TargetMode", "").lower() == "external"
				if rel.get("Target") and not external:
					target = posixpath.normpath(posixpath.join(posixpath.dirname(part), rel["Target"]))
					if rel["Target"].startswith("/"):
						target = rel["Target"].lstrip("/")
				prog_id = obj.get("ProgID", "")
				payload = _payload(archive, names, target)
				is_mathtype = bool(re.fullmatch(r"Equation\.DSMT\d+", prog_id, re.I))
				item = {"part": part, "prog_id": prog_id, "type": obj.get("Type"),
						"relationship_id": rel_id, "target": target, "external": external,
						"declared_mathtype": is_mathtype, **payload}
				ole_objects.append(item)
				if is_mathtype and payload["embedded_exists"] and payload["compound_file"]:
					counts["mathtype_objects"] += 1
				if not payload["embedded_exists"] and not external:
					warnings.append(f"OLE object {prog_id or '(missing ProgID)'} in {part} has no embedded payload")
				elif is_mathtype and not payload["compound_file"]:
					warnings.append(f"MathType declaration in {part} has no Compound File payload")
		# Field codes can contain unresolved internal markers even if not visible.
		for field in fields:
			for match in PLACEHOLDER_RE.finditer(field["code"]):
				placeholders.append({"part": field["part"], "field_code": True, "token": match.group(0)})
		forbidden = [field for field in fields if EQUATION_SEQUENCE_RE.search(field["code"])]
		embeddings = [{"part": name, "size_bytes": archive.getinfo(name).file_size,
					**_payload(archive, names, name)}
					for name in sorted(names) if name.startswith("word/embeddings/") and not name.endswith("/")]
		counts.update({"ole_objects": len(ole_objects), "embeddings": len(embeddings),
					"fields": len(fields), "bookmarks": len(bookmarks), "placeholders": len(placeholders),
					"independent_equation_sequences": len(forbidden)})
		if counts["omml"]:
			warnings.append("OMML formulas are present; this workflow requires TeX or MathType OLE")
		if forbidden:
			warnings.append("Independent SEQ Equation fields are forbidden by the workflow")
		if placeholders:
			warnings.append("Unresolved L2W_ internal placeholders remain")
		styles = _styles(styles_root, samples, warnings)
		known_styles = {style["style_id"] for style in styles}
		unlisted_samples = {key: value for key, value in samples.items() if key not in known_styles}
		return {
			"path": str(path.resolve()), "counts": dict(counts), "warnings": warnings,
			"ole_objects": ole_objects, "embeddings": embeddings, "fields": fields,
			"independent_equation_sequences": forbidden, "bookmarks": bookmarks,
			"placeholders": placeholders, "sections": sections, "styles": styles,
			"document_defaults": _properties(styles_root.find("w:docDefaults", NS)) if styles_root is not None else {},
			"paragraph_style_counts": dict(style_counts), "unlisted_style_samples": unlisted_samples,
			"numbering": _properties(roots.get("word/numbering.xml")),
			"visual_layout_verified": False,
			"mathtype_editability_verified": False,
			"notes": [
				"Style definitions, inheritance chains and direct formatting are reported separately; Word computes effective formatting.",
				"MathType counts require a matching declared ProgID, an existing embedding and a Compound File signature; editability still requires Word/MathType verification.",
				"TeX span counts are heuristic; literal currency/code text may also contain dollar delimiters.",
				"OOXML property values retain native units; page/indent/tab values usually use twips and font sizes half-points.",
			],
		}
