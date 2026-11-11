# author: Codex
"""Apply semantic styles and native Word structures to Pandoc's DOCX.

Math never passes through Pandoc's math writer. Tokens are replaced by literal
TeX and precise bookmarks, which the desktop backend subsequently consumes.
"""
from __future__ import annotations

import copy
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


class WriterError(ValueError):
	pass


def w(name):
	return f"{{{W}}}{name}"


def element(tag, **attrs):
	node = ET.Element(w(tag))
	for key, value in attrs.items():
		node.set(w(key), str(value))
	return node


def child(parent, name):
	node = parent.find(w(name))
	if node is None:
		node = element(name)
		parent.append(node)
	return node


def text_run(text):
	run = element("r")
	t = child(run, "t")
	t.set(XML_SPACE, "preserve")
	t.text = text
	return run


def tab_run():
	run = element("r")
	run.append(element("tab"))
	return run


def field(code, result="0"):
	out = []
	for kind, text in [("begin", None), (None, code), ("separate", None), (None, result), ("end", None)]:
		run = element("r")
		if kind:
			run.append(element("fldChar", fldCharType=kind))
		else:
			node = element("instrText" if len(out) == 1 else "t")
			node.set(XML_SPACE, "preserve")
			node.text = text
			run.append(node)
		out.append(run)
	return out


def replace_token(root, token, replacement):
	"""Replace a Pandoc Str token, retaining surrounding runs and run properties."""
	candidates = root.xpath(".//w:t[contains(text(), $token)]", namespaces=NS, token=token)
	if len(candidates) != 1:
		raise WriterError(f"Expected exactly one token {token!r}; found {len(candidates)}")
	node = candidates[0]
	before, after = node.text.split(token, 1)
	run = node.getparent()
	if run.tag != w("r"):
		raise WriterError(f"Unexpected token container for {token}")
	parent = run.getparent()
	idx = parent.index(run)
	prefix, suffix = copy.deepcopy(run), copy.deepcopy(run)
	target_index = run.index(node)
	for item in list(prefix)[target_index + 1:]:
		prefix.remove(item)
	prefix[target_index].text = before
	for item in list(suffix)[:target_index]:
		if item.tag != w("rPr"):
			suffix.remove(item)
	suffix.find(w("t")).text = after
	parent.remove(run)
	for item in [prefix, *replacement, suffix]:
		parent.insert(idx, item)
		idx += 1
	para = parent if parent.tag == w("p") else next(parent.iterancestors(w("p")), None)
	if para is None:
		raise WriterError(f"Token {token} is not in a paragraph")
	return para


def set_value(parent, name, value):
	child(parent, name).set(w("val"), str(value))


def pstyle(para, style_id):
	pp = para.find(w("pPr"))
	if pp is None:
		pp = element("pPr")
		para.insert(0, pp)
	# pStyle precedes numPr, spacing, etc. in CT_PPr.
	existing = pp.find(w("pStyle"))
	if existing is not None:
		pp.remove(existing)
	pp.insert(0, element("pStyle", val=style_id))
	return pp


def _twips_cm(value):
	return round(float(value) * 1440 / 2.54)


def _style_properties(style, spec, keys):
	rp, pp = child(style, "rPr"), child(style, "pPr")
	for name in keys:
		value = spec[name]
		if name == "latin_font":
			fonts = child(rp, "rFonts")
			for attr in ["ascii", "hAnsi", "cs"]:
				fonts.set(w(attr), value)
				fonts.attrib.pop(w(attr + "Theme"), None)
		elif name == "east_asia_font":
			fonts = child(rp, "rFonts")
			fonts.set(w("eastAsia"), value)
			fonts.attrib.pop(w("eastAsiaTheme"), None)
		elif name == "size_pt":
			set_value(rp, "sz", round(value * 2))
			set_value(rp, "szCs", round(value * 2))
		elif name == "bold":
			set_value(rp, "b", int(value))
		elif name == "alignment":
			set_value(pp, "jc", {"justify": "both"}.get(value, value))
		elif name in ["space_before_pt", "space_after_pt"]:
			child(pp, "spacing").set(w("before" if name == "space_before_pt" else "after"), str(round(value * 20)))
		elif name == "line_spacing":
			spacing = child(pp, "spacing")
			spacing.set(w("line"), str(round(value * 240)))
			spacing.set(w("lineRule"), "auto")
		elif name == "first_line_indent_pt":
			ind = child(pp, "ind")
			for attr in ["hanging", "firstLine", "hangingChars", "firstLineChars"]:
				ind.attrib.pop(w(attr), None)
			ind.set(w("hanging" if value < 0 else "firstLine"), str(round(abs(value) * 20)))
		elif name == "keep_with_next":
			set_value(pp, "keepNext", int(value))
	set_value(pp, "autoSpaceDE", 1)
	set_value(pp, "autoSpaceDN", 1)
	set_value(style, "qFormat", 1)
	# Preserve standard OOXML ordering for Word and other consumers.
	order = {name: i for i, name in enumerate(["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr"])}
	style[:] = sorted(style, key=lambda n: order.get(ET.QName(n).localname, 99))


def _styles(root, config):
	ids = {}
	explicit = config.get("_explicit", {}).get("styles", {})
	has_reference = bool(config.get("reference_docx"))
	defaults = root.xpath("w:style[@w:type='paragraph'][@w:default='1' or @w:default='true']/@w:styleId", namespaces=NS)
	base_style = defaults[0] if defaults else "Normal"
	for semantic, spec in config["styles"].items():
		matches = [st for st in root.findall("w:style", NS)
				if st.get(w("type")) == "paragraph" and st.find(w("name")) is not None
				and st.find(w("name")).get(w("val"), "").casefold() == spec["name"].casefold()]
		fresh = not matches
		if fresh:
			sid = re.sub(r"[^A-Za-z0-9]", "", spec["name"]) or "L2W" + semantic.title()
			existing_ids = set(root.xpath("w:style/@w:styleId", namespaces=NS))
			base, count = sid, 1
			while sid in existing_ids:
				sid = base + str(count)
				count += 1
			style = element("style", type="paragraph", styleId=sid)
			style.append(element("name", val=spec["name"]))
			if sid != base_style:
				style.append(element("basedOn", val=base_style))
			root.append(style)
		else:
			style = matches[0]
			sid = style.get(w("styleId"))
		keys = list(spec) if fresh or not has_reference else explicit.get(semantic, [])
		_style_properties(style, spec, keys)
		if semantic.startswith("heading"):
			set_value(child(style, "pPr"), "outlineLvl", int(semantic[-1]) - 1)
		ids[semantic] = sid
	return ids


def _heading_numbering(styles, numbering, ids, enabled):
	if not enabled:
		return
	by_id = {st.get(w("styleId")): st for st in styles.findall(w("style"))}
	linked = []
	for level in range(1, 4):
		sid = ids[f"heading{level}"]
		st, seen, num_id = by_id[sid], set(), None
		while st is not None and st.get(w("styleId")) not in seen:
			seen.add(st.get(w("styleId")))
			num = st.find("w:pPr/w:numPr/w:numId", NS)
			if num is not None:
				num_id = num.get(w("val"))
				break
			base = st.find(w("basedOn"))
			st = by_id.get(base.get(w("val"))) if base is not None else None
		if num_id in {None, "0"}:
			for abstract in numbering.findall(w("abstractNum")):
				if any(n.get(w("val")) == sid for n in abstract.findall("w:lvl/w:pStyle", NS)):
					candidates = numbering.xpath("w:num[w:abstractNumId/@w:val=$aid]/@w:numId", namespaces=NS, aid=abstract.get(w("abstractNumId")))
					if candidates:
						num_id = candidates[0]
						break
		linked.append(num_id if num_id != "0" else None)
	if len({n for n in linked if n}) > 1:
		raise WriterError("Template heading numbering uses inconsistent sequences; provide a coherent three-level list")
	if all(linked):
		# Existing style/list association and numbering appearance remain intact.
		return
	if any(linked):
		existing = next(n for n in linked if n)
		if any(n and n != existing for n in linked):
			raise WriterError("Template heading numbering uses inconsistent sequences; provide a coherent three-level list")
		matches = numbering.xpath("w:num[@w:numId=$nid]/w:abstractNumId/@w:val", namespaces=NS, nid=existing)
		abstracts = numbering.xpath("w:abstractNum[@w:abstractNumId=$aid]", namespaces=NS, aid=matches[0]) if matches else []
		if not abstracts:
			raise WriterError("Template heading numbering references a missing list definition")
		abstract = abstracts[0]
		for index, current in enumerate(linked):
			if current:
				continue
			sid = ids[f"heading{index + 1}"]
			levels = abstract.xpath("w:lvl[@w:ilvl=$level]", namespaces=NS, level=str(index))
			if levels:
				lvl = levels[0]
			else:
				lvl = element("lvl", ilvl=index)
				for name, value in [("start", 1), ("numFmt", "decimal"), ("lvlText", ".".join(f"%{i + 1}" for i in range(index + 1))), ("lvlJc", "left")]:
					lvl.append(element(name, val=value))
				abstract.append(lvl)
			set_value(lvl, "pStyle", sid)
			np = child(child(by_id[sid], "pPr"), "numPr")
			set_value(np, "ilvl", index)
			set_value(np, "numId", existing)
		return
	abs_ids = [int(x) for x in numbering.xpath("w:abstractNum/@w:abstractNumId", namespaces=NS)]
	num_ids = [int(x) for x in numbering.xpath("w:num/@w:numId", namespaces=NS)]
	aid, nid = max(abs_ids + [-1]) + 1, max(num_ids + [0]) + 1
	abstract = element("abstractNum", abstractNumId=aid)
	abstract.append(element("multiLevelType", val="multilevel"))
	for level in range(3):
		lvl = element("lvl", ilvl=level)
		for name, value in [("start", 1), ("numFmt", "decimal"), ("pStyle", ids[f"heading{level + 1}"]), ("suff", "space"), ("lvlText", ".".join(f"%{i + 1}" for i in range(level + 1))), ("lvlJc", "left")]:
			lvl.append(element(name, val=value))
		abstract.append(lvl)
		st = styles.xpath("w:style[@w:styleId=$sid]", namespaces=NS, sid=ids[f"heading{level + 1}"])[0]
		np = child(child(st, "pPr"), "numPr")
		set_value(np, "ilvl", level)
		set_value(np, "numId", nid)
	numbering.insert(len(numbering.findall(w("abstractNum"))), abstract)
	num = element("num", numId=nid)
	num.append(element("abstractNumId", val=aid))
	numbering.append(num)


def _pages(root, config):
	spec = config["page"]
	keys = list(spec) if not config.get("reference_docx") else config.get("_explicit", {}).get("page", [])
	sections = root.findall(".//w:sectPr", NS)
	if not sections:
		sections = [child(root.find(w("body")), "sectPr")]
	for section in sections:
		size, margin = child(section, "pgSz"), child(section, "pgMar")
		for key in keys:
			if key in ["width_cm", "height_cm"]:
				size.set(w("w" if key == "width_cm" else "h"), str(_twips_cm(spec[key])))
			elif key.startswith("margin_"):
				margin.set(w(key.split("_")[1]), str(_twips_cm(spec[key])))
	return sections


def _equation_layout(para, sections, numbered, style_id, styles):
	if next(para.iterancestors(w("tc")), None) is not None:
		raise WriterError("Display equations inside table cells require container-specific layout; use inline math")
	section = sections[-1]
	# Pandoc writes one section. Explicitly reject unsupported multi-section layouts.
	if len(sections) != 1:
		raise WriterError("Display equations in multiple sections are not supported by this layout profile")
	cols = section.find(w("cols"))
	if cols is not None and (int(cols.get(w("num"), "1")) > 1 or len(cols.findall(w("col"))) > 1):
		raise WriterError("Display equations require a single-column reference document")
	size, margin = section.find(w("pgSz")), section.find(w("pgMar"))
	width = int(size.get(w("w"))) - sum(int(margin.get(w(k), "0")) for k in ["left", "right", "gutter"])
	pp = pstyle(para, style_id)
	ind = child(pp, "ind")
	ind.attrib.clear()
	for key in ["left", "right", "firstLine", "leftChars", "rightChars", "firstLineChars"]:
		ind.set(w(key), "0")
	set_value(pp, "jc", "left")
	set_value(pp, "keepNext", 0)
	spacing = child(pp, "spacing")
	spacing.set(w("line"), "240")
	spacing.set(w("lineRule"), "auto")
	inherited_positions = set()
	by_id = {st.get(w("styleId")): st for st in styles.findall(w("style"))}
	st, seen = by_id.get(style_id), set()
	while st is not None and st.get(w("styleId")) not in seen:
		seen.add(st.get(w("styleId")))
		inherited_positions.update(int(n.get(w("pos"))) for n in st.findall("w:pPr/w:tabs/w:tab", NS))
		base = st.find(w("basedOn"))
		st = by_id.get(base.get(w("val"))) if base is not None else None
	inherited_positions.update(int(n.get(w("pos"))) for n in pp.findall("w:tabs/w:tab", NS))
	tabs = child(pp, "tabs")
	tabs.clear()
	# Clear inherited tab stops before introducing the two intentional positions.
	for pos in sorted(inherited_positions - {round(width / 2), width}):
		tabs.append(element("tab", val="clear", pos=pos))
	tabs.append(element("tab", val="center", pos=round(width / 2)))
	if numbered:
		tabs.append(element("tab", val="right", pos=width))
	# Do not rely on inherited first-line or list formatting.
	np = pp.find(w("numPr"))
	if np is not None:
		pp.remove(np)


def style_docx(input_path: Path, output_path: Path, manifest: dict, config: dict) -> dict:
	"""Write a TeX work document; native MathType mutation is a later stage."""
	with ZipFile(input_path) as archive:
		files = {name: archive.read(name) for name in archive.namelist()}
	parser = ET.XMLParser(resolve_entities=False, no_network=True)
	roots = {name: ET.fromstring(data, parser) for name, data in files.items()
			if name.startswith(("word/", "docProps/")) and name.endswith(".xml")}
	document = roots["word/document.xml"]
	styles = roots["word/styles.xml"]
	numbering = roots.get("word/numbering.xml")
	if numbering is None:
		raise WriterError("Pandoc did not supply word/numbering.xml")
	ids = _styles(styles, config)
	_heading_numbering(styles, numbering, ids, config["headings"]["numbered"])
	sections = _pages(document, config)
	old_to_new = {"Normal": "body", "BodyText": "body", "FirstParagraph": "body", "Title": "title", "Caption": "caption", "ImageCaption": "figure_caption", "TableCaption": "table_caption", "SourceCode": "code", "Figure": "figure", "CaptionedFigure": "figure", "Heading1": "heading1", "Heading2": "heading2", "Heading3": "heading3"}
	content_roots = [root for name, root in roots.items() if re.match(r"word/(document|footnotes|endnotes|header\d+|footer\d+)\.xml$", name)]
	for root in content_roots:
		for para in root.findall(".//w:p", NS):
			current = para.find("w:pPr/w:pStyle", NS)
			sid = current.get(w("val")) if current is not None else "Normal"
			semantic = old_to_new.get(sid)
			if next(para.iterancestors(w("tc")), None) is not None:
				semantic = "table"
			if semantic:
				pstyle(para, ids[semantic])
	bookmark_ids = [int(v) for root in roots.values() for v in root.xpath(".//w:bookmarkStart/@w:id", namespaces=NS)]
	next_id = max(bookmark_ids + [0]) + 1

	def marked(name, nodes):
		nonlocal next_id
		result = [element("bookmarkStart", id=next_id, name=name), *nodes, element("bookmarkEnd", id=next_id)]
		next_id += 1
		return result

	def replace(token, nodes):
		matching = [root for root in content_roots if root.xpath(".//w:t[contains(text(), $token)]", namespaces=NS, token=token)]
		if len(matching) != 1:
			raise WriterError(f"Token missing or duplicated across document parts: {token}")
		return replace_token(matching[0], token, nodes)

	alternative_text = {}
	for eq in manifest.get("equations", []):
		alternative_text[eq["token"]] = "$" + eq["tex"] + "$"
		nodes = marked(eq["bookmark"], [text_run("$" + eq["tex"] + "$")])
		if eq["display"]:
			nodes.insert(0, tab_run())
			if eq["numbered"]:
				nodes.extend([tab_run(), text_run("L2W_NUMBER_" + eq["id"])])
		para = replace(eq["token"], nodes)
		if eq["display"]:
			_equation_layout(para, sections, eq["numbered"], ids["equation"], styles)
	counters = {"figure": 0, "table": 0}
	label_results = {}
	for target in manifest.get("targets", []):
		kind = target["kind"]
		if kind in counters:
			counters[kind] += 1
			label = config["captions"][kind + "_label"]
			alternative_text[target["token"]] = label + " " + str(counters[kind])
			nodes = [text_run(label + " "), *marked(target["bookmark"], field(f" SEQ {'Figure' if kind == 'figure' else 'Table'} \\* ARABIC ", str(counters[kind])))]
			para = replace(target["token"], nodes)
			pstyle(para, ids[kind + "_caption"])
			label_results[target.get("label")] = str(counters[kind])
		elif kind == "heading":
			alternative_text[target["token"]] = ""
			para = replace(target["token"], [])
			nodes = marked(target["bookmark"], [])
			para.insert(1 if para.find(w("pPr")) is not None else 0, nodes[0])
			para.append(nodes[-1])
			if not target.get("numbered", True) or not config["headings"]["numbered"]:
				set_value(child(pstyle(para, ids[f"heading{target['level']}"]), "numPr"), "numId", 0)
	targets = {t["label"]: t for t in manifest.get("targets", []) if t.get("label")}
	for ref in manifest.get("references", []):
		alternative_text[ref["token"]] = label_results.get(ref["target"], ref["target"])
		if ref["kind"] == "equation":
			replace(ref["token"], marked(ref["bookmark"], [text_run(ref["token"])]))
		else:
			if ref["target"] not in targets:
				raise WriterError(f"Unknown reference target: {ref['target']}")
			target = targets[ref["target"]]
			switch = " \\n" if target["kind"] == "heading" else ""
			replace(ref["token"], marked(ref["bookmark"], field(f" REF {target['bookmark']}{switch} \\h ", label_results.get(ref["target"], "0"))))
	for root in content_roots:
		# Pandoc also copies captions to table/drawing accessibility attributes.
		# Those alternatives remain descriptive text, not live Word fields.
		for node in root.iter():
			for key, value in list(node.attrib.items()):
				if "L2W_" in value:
					for token, readable in alternative_text.items():
						value = value.replace(token, readable)
					node.set(key, value)
		if root.xpath(".//*[local-name()='oMath' or local-name()='oMathPara']"):
			raise WriterError("Unexpected OMML detected; TeX provenance invariant violated")
	for name, root in roots.items():
		if name.startswith("docProps/"):
			for node in root.iter():
				if node.text and "L2W_" in node.text:
					for token, readable in alternative_text.items():
						node.text = node.text.replace(token, readable)
	settings = roots.get("word/settings.xml")
	if settings is not None:
		set_value(settings, "updateFields", "true")
	orders = {
		"pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange",
		"rPr": "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange",
		"sectPr": "headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange",
		"lvl": "start numFmt lvlRestart pStyle isLgl suff lvlText lvlPicBulletId legacy lvlJc pPr rPr",
	}
	for name, root in roots.items():
		for tag, order in orders.items():
			positions = {key: i for i, key in enumerate(order.split())}
			for node in root.iter(w(tag)):
				node[:] = sorted(node, key=lambda n: positions.get(ET.QName(n).localname, 999))
		files[name] = ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
	output_path.parent.mkdir(parents=True, exist_ok=True)
	with ZipFile(output_path, "w", ZIP_DEFLATED) as archive:
		for name, data in files.items():
			archive.writestr(name, data)
	return {"styles": ids, "equations": len(manifest.get("equations", [])), "layout": "properties_set; Word visual verification required"}
