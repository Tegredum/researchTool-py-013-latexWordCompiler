# author: Codex
"""LaTeX -> Pandoc AST, auditable Markdown, and semantic placeholders.

Pandoc parses the document.  The small lexical pass in this module only expands
confined local ``input`` files and annotates math: Pandoc intentionally forgets
whether display math came from ``equation`` or ``equation*``.  Formulae never
reach the DOCX writer as Pandoc Math nodes, so that writer cannot create OMML.
"""

from __future__ import annotations

import copy
import json
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlsplit


class FrontendError(ValueError):
	"""Source cannot be converted without hiding a semantic loss."""


_COMMAND = re.compile(r"\\([A-Za-z@]+|[^\n])")
_MARKER = re.compile(r"\\LTwoWMathMarker\{(s\d+)\}")
_RESERVED = re.compile(r"L2W_|\\LTwoWMathMarker\b")
_VERBATIM = {"verbatim", "Verbatim", "lstlisting", "minted"}
_MATH_ENVS = {"equation", "equation*", "displaymath", "math"}
_SPLIT_ENVS = {"align", "align*", "gather", "gather*"}
_UNSUPPORTED_ENVS = {
	"minipage", "multicols", "multicols*", "subfigure", "subtable",
	"wrapfigure", "wraptable", "tikzpicture", "picture", "pspicture",
	"multline", "multline*", "flalign", "flalign*", "alignat", "alignat*",
	"eqnarray", "eqnarray*", "split", "aligned", "alignedat", "gathered",
}
_DANGEROUS_COMMANDS = {
	"write", "openout", "openin", "read", "catcode", "csname", "directlua",
	"special", "immediate", "scantokens", "includeonly", "endlinechar",
	"usepackagewithoptions", "InputIfFileExists", "IfFileExists", "includegraphics*",
}
_SUPPORTED_IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".svg", ".emf", ".wmf"}


def _location(text: str, offset: int, source: Path | str) -> str:
	line = text.count("\n", 0, offset) + 1
	return f"{source}:{line}"


def _group(text: str, position: int, *, context: str) -> tuple[str, int]:
	"""Read a balanced brace argument, respecting escaped braces and comments."""
	while position < len(text) and text[position].isspace():
		position += 1
	if position >= len(text) or text[position] != "{":
		raise FrontendError(f"{context}: expected a literal {{...}} argument")
	start = position + 1
	depth = 1
	position += 1
	while position < len(text):
		char = text[position]
		if char == "\\":
			match = _COMMAND.match(text, position)
			position = match.end() if match else position + 1
			continue
		if char == "%":
			end = text.find("\n", position)
			position = len(text) if end < 0 else end + 1
			continue
		if char == "{":
			depth += 1
		elif char == "}":
			depth -= 1
			if depth == 0:
				return text[start:position], position + 1
		position += 1
	raise FrontendError(f"{context}: unclosed brace argument")


def _verbatim_end(text: str, position: int, command: re.Match[str], context: str) -> int | None:
	if command.group(1) == "verb":
		start = command.end()
		if start < len(text) and text[start] == "*":
			start += 1
		if start >= len(text) or text[start].isspace():
			raise FrontendError(f"{context}: invalid \\verb delimiter")
		end = text.find(text[start], start + 1)
		if end < 0:
			raise FrontendError(f"{context}: unclosed \\verb")
		return end + 1
	if command.group(1) == "begin":
		name, begin_end = _group(text, command.end(), context=context)
		if name in _VERBATIM:
			end_marker = "\\end{" + name + "}"
			end = text.find(end_marker, begin_end)
			if end < 0:
				raise FrontendError(f"{context}: unclosed {name} environment")
			return end + len(end_marker)
	return None


def _local_path(value: str, root: Path, *, description: str, extensions: bool = False) -> Path:
	value = unquote(value.strip())
	if not value or any(character in value for character in ("\x00", "\n", "\r", "{", "}", "#")):
		raise FrontendError(f"{description}: expected a literal local path, got {value!r}")
	if value.startswith(("//", "\\\\")) or (urlsplit(value).scheme and not re.match(r"^[A-Za-z]:[/\\]", value)):
		raise FrontendError(f"{description}: network/URI resources are not allowed: {value!r}")
	candidate = Path(value)
	if not candidate.is_absolute():
		candidate = root / candidate
	if extensions and not candidate.suffix:
		candidate = candidate.with_suffix(".tex")
	candidate = candidate.resolve()
	if not candidate.is_relative_to(root):
		raise FrontendError(f"{description}: path escapes source directory {root}: {candidate}")
	if not candidate.is_file():
		raise FrontendError(f"{description}: local file not found: {candidate}")
	return candidate


def _expand_inputs(path: Path, root: Path, sources: list[str], stack: tuple[Path, ...] = ()) -> str:
	if path in stack:
		raise FrontendError("Recursive LaTeX input: " + " -> ".join(str(item) for item in (*stack, path)))
	if len(stack) >= 32:
		raise FrontendError("LaTeX input nesting exceeds 32 files")
	try:
		text = path.read_text(encoding="utf-8-sig")
	except UnicodeError as error:
		raise FrontendError(f"{path}: source must be UTF-8: {error}") from error
	if _RESERVED.search(text):
		raise FrontendError(f"{path}: source uses a reserved L2W placeholder/marker")
	if str(path) not in sources:
		sources.append(str(path))
	output: list[str] = []
	position = 0
	while position < len(text):
		char = text[position]
		if char == "%":
			end = text.find("\n", position)
			end = len(text) if end < 0 else end + 1
			output.append(text[position:end])
			position = end
			continue
		command = _COMMAND.match(text, position) if char == "\\" else None
		if not command:
			output.append(char)
			position += 1
			continue
		context = _location(text, position, path)
		protected_end = _verbatim_end(text, position, command, context)
		if protected_end is not None:
			output.append(text[position:protected_end])
			position = protected_end
			continue
		name = command.group(1)
		if name in _DANGEROUS_COMMANDS:
			raise FrontendError(f"{context}: unsupported execution/dynamic-I/O command \\{name}")
		if name in {"input", "include"}:
			start = command.end()
			while start < len(text) and text[start].isspace():
				start += 1
			if start < len(text) and text[start] == "{":
				filename, end = _group(text, start, context=context)
			else:
				argument = re.match(r"[^\s%{}]+", text[start:])
				if not argument:
					raise FrontendError(f"{context}: \\{name} needs a literal file name")
				filename, end = argument.group(), start + argument.end()
			if "\\" in filename or "#" in filename:
				raise FrontendError(f"{context}: dynamic LaTeX input paths are unsupported")
			included = _local_path(filename, root, description=context + f" \\{name}", extensions=True)
			output.append(_expand_inputs(included, root, sources, (*stack, path)))
			position = end
			continue
		output.append(command.group())
		position = command.end()
	return "".join(output)


def _environment_end(text: str, start: int, name: str, context: str) -> tuple[str, int]:
	"""Locate an environment's matched end; brace groups do not hide environments."""
	stack = [name]
	position = start
	while position < len(text):
		if text[position] == "%":
			end = text.find("\n", position)
			position = len(text) if end < 0 else end + 1
			continue
		match = _COMMAND.match(text, position)
		if not match:
			position += 1
			continue
		if match.group(1) in {"begin", "end"}:
			environment, end = _group(text, match.end(), context=context)
			if match.group(1) == "begin":
				stack.append(environment)
			else:
				if not stack or stack[-1] != environment:
					raise FrontendError(f"{context}: mismatched \\end{{{environment}}}")
				stack.pop()
				if not stack:
					return text[start:position], end
			position = end
		else:
			position = match.end()
	raise FrontendError(f"{context}: unclosed {name} environment")


def _delimited_math(text: str, start: int, delimiter: str, context: str) -> tuple[str, int]:
	position = start
	depth = 0
	while position < len(text):
		if depth == 0 and text.startswith(delimiter, position):
			return text[start:position], position + len(delimiter)
		if text[position] == "%":
			end = text.find("\n", position)
			position = len(text) if end < 0 else end + 1
			continue
		match = _COMMAND.match(text, position)
		if not match:
			if text[position] == "{":
				depth += 1
			elif text[position] == "}":
				depth -= 1
		position = match.end() if match else position + 1
	raise FrontendError(f"{context}: unclosed math delimiter {delimiter!r}")


def _labels(tex: str, context: str) -> tuple[str, list[str]]:
	parts: list[str] = []
	labels: list[str] = []
	position = 0
	while position < len(tex):
		match = _COMMAND.match(tex, position)
		if match and match.group(1) == "label":
			label, end = _group(tex, match.end(), context=context)
			if not label or re.search(r"[\s{}\\]", label):
				raise FrontendError(f"{context}: unsupported label {label!r}")
			labels.append(label)
			position = end
		elif match:
			parts.append(match.group())
			position = match.end()
		else:
			parts.append(tex[position])
			position += 1
	return "".join(parts), labels


def _without_comments(text: str) -> str:
	output: list[str] = []
	position = 0
	while position < len(text):
		command = _COMMAND.match(text, position)
		if command:
			output.append(command.group())
			position = command.end()
		elif text[position] == "%":
			end = text.find("\n", position)
			position = len(text) if end < 0 else end + 1
		else:
			output.append(text[position])
			position += 1
	return "".join(output)


def _check_math(tex: str, context: str) -> None:
	tex = _without_comments(tex)
	if re.search(r"\\(?:tag\*?|eqno|leqno|notag|nonumber)\b", tex):
		raise FrontendError(f"{context}: explicit tags or per-row number suppression are unsupported; use equation/equation* with numbering outside TeX")
	for match in re.finditer(r"\\begin\s*\{([^}]+)\}", tex):
		if match.group(1) in _UNSUPPORTED_ENVS | _SPLIT_ENVS:
			raise FrontendError(f"{context}: nested {match.group(1)} cannot be safely reduced to independent single-line equations")


def _split_rows(body: str, name: str, context: str) -> list[str]:
	"""Conservatively split independent align/gather rows, never continuations."""
	body = _without_comments(body)
	if re.search(r"\\(?:begin|intertext|shortintertext|notag|nonumber|tag)\b", body):
		raise FrontendError(f"{context}: {name} has nested layout/text/number overrides; rewrite it as independent equation environments")
	rows: list[str] = []
	start = position = depth = 0
	while position < len(body):
		match = _COMMAND.match(body, position)
		if match:
			if match.group(1) == "\\" and depth == 0:
				rows.append(body[start:position].strip())
				position = match.end()
				if position < len(body) and body[position] in "[*":
					raise FrontendError(f"{context}: row spacing/star overrides in {name} are unsupported")
				start = position
				continue
			position = match.end()
			continue
		if body[position] == "{":
			depth += 1
		elif body[position] == "}":
			depth -= 1
		position += 1
	rows.append(body[start:].strip())
	if rows and not rows[-1]:
		rows.pop()
	result: list[str] = []
	for row in rows:
		stripped, _ = _labels(row, context)
		if name.startswith("align"):
			# Multiple columns and continuation rows have grouping semantics.
			if stripped.count("&") != 1:
				raise FrontendError(f"{context}: {name} requires exactly one alignment point in each independent row")
			left, right = stripped.split("&")
			relation = re.match(r"\s*(?:=|\\(?:approx|equiv|leq?|geq?|sim|simeq|to|in|propto|neq)\b)\s*(\S[\s\S]*)", right)
			if not left.strip() or not relation or re.match(r"\s*(?:[=+\-]|\\(?:approx|equiv|leq?|geq?))", left):
				raise FrontendError(f"{context}: continuation or ambiguous {name} row; use independent complete equations")
			row = row.replace("&", "")
		elif "&" in stripped or not stripped.strip() or re.match(r"\s*[=+\-]", stripped):
			raise FrontendError(f"{context}: ambiguous {name} row cannot be split safely")
		result.append(row.strip())
	if not result:
		raise FrontendError(f"{context}: empty {name} environment")
	return result


def _annotate_math(text: str, source: Path, warnings: list[str]) -> tuple[str, dict[str, dict[str, Any]]]:
	descriptors: dict[str, dict[str, Any]] = {}
	output: list[str] = []
	position = 0

	def annotated(body: str, display: bool, numbered: bool, context: str) -> str:
		body = _without_comments(body)
		_check_math(body, context)
		marker = f"s{len(descriptors) + 1:06d}"
		descriptors[marker] = {"display": display, "numbered": numbered, "location": context}
		return "\\LTwoWMathMarker{" + marker + "}" + body

	while position < len(text):
		context = _location(text, position, f"{source} (expanded source)")
		if text[position] == "%":
			end = text.find("\n", position)
			end = len(text) if end < 0 else end + 1
			output.append(text[position:end])
			position = end
			continue
		if text[position] == "$":
			delimiter = "$$" if text.startswith("$$", position) else "$"
			body, end = _delimited_math(text, position + len(delimiter), delimiter, context)
			output.append(delimiter + annotated(body, len(delimiter) == 2, False, context) + delimiter)
			position = end
			continue
		command = _COMMAND.match(text, position)
		if not command:
			output.append(text[position])
			position += 1
			continue
		protected_end = _verbatim_end(text, position, command, context)
		if protected_end is not None:
			output.append(text[position:protected_end])
			position = protected_end
			continue
		if command.group(1) in {"[", "("}:
			closing = "\\]" if command.group(1) == "[" else "\\)"
			body, end = _delimited_math(text, command.end(), closing, context)
			output.append(command.group() + annotated(body, closing == "\\]", False, context) + closing)
			position = end
			continue
		if command.group(1) == "ensuremath":
			body, end = _group(text, command.end(), context=context)
			output.append("\\ensuremath{" + annotated(body, False, False, context) + "}")
			position = end
			continue
		if command.group(1) == "begin":
			name, start = _group(text, command.end(), context=context)
			if name in _UNSUPPORTED_ENVS:
				raise FrontendError(f"{context}: unsupported layout environment {name}; supply a normal paragraph/table/image or independent equation")
			if name in _MATH_ENVS | _SPLIT_ENVS:
				body, end = _environment_end(text, start, name, context)
				if name in _SPLIT_ENVS:
					rows = _split_rows(body, name, context)
					numbered = not name.endswith("*")
					environment = "equation" if numbered else "equation*"
					output.append("\n\n".join("\\begin{" + environment + "}" + annotated(row, True, numbered, context) + "\\end{" + environment + "}" for row in rows))
					warnings.append(f"{context}: normalized {name} into {len(rows)} independent display equations; review the original grouping")
				else:
					output.append("\\begin{" + name + "}" + annotated(body, name != "math", name == "equation", context) + "\\end{" + name + "}")
				position = end
				continue
		output.append(command.group())
		position = command.end()
	return "".join(output), descriptors


def find_pandoc(executable: str | None = None) -> str:
	"""Resolve an explicitly supplied Pandoc or the executable on PATH."""
	found = shutil.which(executable or "pandoc")
	if not found and executable and Path(executable).is_file():
		found = str(Path(executable).resolve())
	if not found:
		raise FrontendError("Pandoc was not found; install Pandoc 3.x or supply --pandoc with its executable path")
	return str(Path(found).resolve())


def _pandoc(executable: str, arguments: list[str], text: str, cwd: Path) -> tuple[str, list[str]]:
	try:
		result = subprocess.run(
			[executable, *arguments], input=text, encoding="utf-8", errors="strict",
			stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=cwd, timeout=90,
			check=False, shell=False,
		)
	except (OSError, subprocess.TimeoutExpired) as error:
		raise FrontendError(f"Pandoc could not complete: {error}") from error
	if result.returncode:
		raise FrontendError(f"Pandoc failed (exit {result.returncode}): {result.stderr.strip()}")
	return result.stdout, [line for line in result.stderr.splitlines() if line.strip()]


def _plain(value: Any) -> str:
	if isinstance(value, list):
		return "".join(_plain(item) for item in value)
	if not isinstance(value, dict):
		return ""
	kind = value.get("t")
	if kind == "Str":
		return value["c"]
	if kind in {"Space", "SoftBreak", "LineBreak"}:
		return " "
	if kind in {"Code", "Math"}:
		return value["c"][1]
	if kind in {"Link", "Image", "Span"}:
		return _plain(value["c"][1])
	return _plain(value.get("c", []))


def _nodes(value: Any):
	if isinstance(value, dict):
		if "t" in value:
			yield value
		for child in value.values():
			yield from _nodes(child)
	elif isinstance(value, list):
		for child in value:
			yield from _nodes(child)


def _bibliography(ast: dict[str, Any], root: Path, executable: str, descriptors: dict[str, dict[str, Any]], warnings: list[str]) -> tuple[dict[str, Any], list[str], list[str]]:
	"""Resolve confined .bib files through bundled citeproc, never BibTeX/TeX."""
	metadata = ast.get("meta", {})
	citations = sorted({citation["citationId"] for node in _nodes(ast) if node.get("t") == "Cite" for citation in node["c"][0]})
	has_print = any(node.get("t") in {"RawInline", "RawBlock"} and node["c"][1].strip() == r"\printbibliography" for node in _nodes(ast))
	bibliography = metadata.get("bibliography")
	if bibliography is None:
		if citations or has_print:
			raise FrontendError("Citations/\\printbibliography need a local .bib file declared by \\bibliography or \\addbibresource")
		return ast, [], []
	if any(key in metadata for key in ("csl", "citation-abbreviations")):
		raise FrontendError("Source CSL/abbreviation resources are unsupported; this converter uses Pandoc's default CSL style")

	def names(value: dict[str, Any]) -> list[str]:
		if value["t"] == "MetaList":
			return [item for child in value["c"] for item in names(child)]
		if value["t"] == "MetaString":
			return [value["c"]]
		if value["t"] == "MetaInlines":
			return [_plain(value["c"])]
		raise FrontendError("Unsupported bibliography metadata; use literal .bib file names")

	files = [_local_path(name, root, description="Bibliography") for name in names(bibliography)]
	if not files or any(path.suffix.lower() != ".bib" for path in files):
		raise FrontendError("Bibliography inputs must be local .bib files")
	for path in files:
		try:
			content = path.read_text(encoding="utf-8-sig")
		except UnicodeError as error:
			raise FrontendError(f"{path}: bibliography must be UTF-8: {error}") from error
		if _RESERVED.search(content):
			raise FrontendError(f"{path}: bibliography uses a reserved L2W placeholder/marker")
	for node in _nodes(ast):
		if node.get("t") == "Math" and not _MARKER.search(node["c"][1]):
			raise FrontendError(f"Cannot recover source numbering semantics for formula {node['c'][1][:120]!r}")

	dropped = object()

	def remove_print(value: Any) -> Any:
		if isinstance(value, list):
			return [processed for child in value if (processed := remove_print(child)) is not dropped]
		if not isinstance(value, dict):
			return value
		if value.get("t") in {"RawInline", "RawBlock"} and value["c"][1].strip() == r"\printbibliography":
			return dropped
		return {key: remove_print(child) for key, child in value.items()}

	ast = remove_print(ast)
	# Do not let a relative metadata path trigger a second read outside the
	# explicit, validated command-line resources.
	ast["meta"].pop("bibliography", None)
	output, messages = _pandoc(executable, ["--from=json", "--to=json", "--sandbox", "--citeproc", *["--bibliography=" + str(path) for path in files]], json.dumps(ast, ensure_ascii=False), root)
	if messages:
		raise FrontendError("Citeproc reported unresolved/invalid bibliography data: " + " | ".join(messages))
	ast = json.loads(output)
	ast["meta"].pop("bibliography", None)
	ast["meta"].pop("nocite", None)
	for node in _nodes(ast):
		if node.get("t") == "Math" and not _MARKER.search(node["c"][1]):
			if node["c"][0]["t"] != "InlineMath":
				raise FrontendError("Display math introduced by a bibliography entry is unsupported")
			marker = f"s{len(descriptors) + 1:06d}"
			descriptors[marker] = {"display": False, "numbered": False, "location": "bibliography/citeproc"}
			node["c"][1] = "\\LTwoWMathMarker{" + marker + "}" + node["c"][1]
		elif node.get("t") == "Cite":
			# Freeze rendered citation text so Markdown and DOCX show exactly
			# the same citation rather than Markdown re-emitting @keys.
			node["t"], node["c"] = "Span", [["", ["citation"], []], node["c"][1]]
	warnings.append("Bibliography and citations are rendered with Pandoc's default CSL style as editable text; original .bst/biblatex styling and dynamic Word bibliography fields are not reproduced")
	if has_print:
		warnings.append("\\printbibliography is rendered at the end of the document; custom bibliography headings/placement options are unsupported")
	return ast, [str(path) for path in files], citations


class _Transformer:
	def __init__(self, root: Path, descriptors: dict[str, dict[str, Any]], strict: bool, warnings: list[str], source: Path, source_text: str):
		self.root, self.descriptors, self.strict, self.warnings = root, descriptors, strict, warnings
		self.source, self.source_text = source, source_text
		self.equations: list[dict[str, Any]] = []
		self.references: list[dict[str, Any]] = []
		self.targets: list[dict[str, Any]] = []
		self.assets: list[str] = []
		self.labels: dict[str, dict[str, Any]] = {}
		self.counts = {"heading": 0, "figure": 0, "table": 0}

	def register(self, item: dict[str, Any]) -> None:
		label = item.get("label")
		if label:
			if label in self.labels:
				raise FrontendError(f"Duplicate semantic label {label!r}; every equation/figure/table/heading needs a unique label")
			self.labels[label] = item

	def target(self, kind: str, label: str, caption: Any, *, level: int | None = None, numbered: bool = True) -> dict[str, Any]:
		self.counts[kind] += 1
		prefix, bookmark = {"heading": ("sec", "Sec"), "figure": ("fig", "Fig"), "table": ("tbl", "Tbl")}[kind]
		identifier = f"{prefix}{self.counts[kind]:04d}"
		item: dict[str, Any] = {
			"id": identifier, "kind": kind, "label": label or None,
			"bookmark": f"L2W{bookmark}{self.counts[kind]:04d}",
			"token": "L2W_CAP_" + identifier, "caption": _plain(caption), "numbered": numbered,
		}
		if level is not None:
			item["level"] = level
		self.targets.append(item)
		self.register(item)
		return item

	def caption(self, node: dict[str, Any], kind: str) -> None:
		attr, caption = node["c"][0:2]
		short = caption[0]
		if short and any(item.get("t") in {"Math", "RawInline", "RawBlock", "Cite"} for item in _nodes(short)):
			raise FrontendError(f"Math/commands/references in a short {kind} caption are unsupported; put them in the long body caption")
		blocks = caption[1]
		if not blocks:
			if attr[0]:
				raise FrontendError(f"{kind} label {attr[0]!r} has no caption and therefore no numbered target")
			return
		if blocks[0]["t"] not in {"Para", "Plain"}:
			raise FrontendError(f"Unsupported complex {kind} caption: first caption block must be text")
		target = self.target(kind, attr[0], blocks)
		if short:
			target["short_caption"] = _plain(short)
			self.warnings.append(f"Short {kind} caption is retained in the manifest/Markdown; Word uses the long caption and does not generate a separate list of figures/tables")
		blocks[0]["c"][0:0] = [{"t": "Str", "c": target["token"]}, {"t": "Space"}]

	def reference(self, command: str, label: str, in_note: bool) -> dict[str, Any]:
		if in_note:
			raise FrontendError("Cross-references in footnotes/endnotes are not supported by the Word main-story conversion")
		if command not in {"ref", "eqref"}:
			raise FrontendError(f"Unsupported cross-reference command \\{command}; use \\ref or \\eqref")
		number = len(self.references) + 1
		item = {"token": f"L2W_REF_{number:04d}", "bookmark": f"L2WRef{number:04d}", "target": label, "kind": None, "command": command}
		self.references.append(item)
		return {"t": "Str", "c": item["token"]}

	def walk(self, value: Any, *, in_note: bool = False) -> Any:
		if isinstance(value, list):
			transformed = [self.walk(item, in_note=in_note) for item in value]
			return [item for item in transformed if not (isinstance(item, dict) and item.get("t") == "Null")]
		if not isinstance(value, dict):
			return value
		node = copy.deepcopy(value)
		kind = node.get("t")
		if kind == "Note":
			in_note = True
		if kind == "Div":
			attr, blocks = node["c"]
			if attr[0] and len(blocks) == 1 and blocks[0].get("t") == "Table":
				blocks[0]["c"][0][0] = attr[0]
				attr[0] = ""
		if kind == "Header":
			level, attr, inlines = node["c"]
			if not 1 <= level <= 3:
				raise FrontendError(f"Heading level {level} exceeds the workflow's three-level limit: {_plain(inlines)!r}")
			target = self.target("heading", attr[0], inlines, level=level, numbered="unnumbered" not in attr[1])
			inlines.insert(0, {"t": "Str", "c": target["token"]})
		elif kind in {"Figure", "Table"}:
			self.caption(node, "figure" if kind == "Figure" else "table")
		elif kind == "Image":
			filename = node["c"][2][0]
			image_path = _local_path(filename, self.root, description="Image " + repr(filename))
			if image_path.suffix.lower() not in _SUPPORTED_IMAGES:
				raise FrontendError(f"Image {filename!r}: convert the local asset to a Word-supported image (e.g. PNG/SVG) first; keep its generation source")
			node["c"][2][0] = image_path.as_posix()
			if str(image_path) not in self.assets:
				self.assets.append(str(image_path))
		elif kind == "Math":
			if in_note:
				raise FrontendError("Math in a footnote/endnote is not supported by the Word main-story conversion; move it into the body")
			display = node["c"][0]["t"] == "DisplayMath"
			tex = node["c"][1]
			markers = _MARKER.findall(tex)
			if len(markers) != 1 or markers[0] not in self.descriptors:
				raise FrontendError(f"Cannot recover source numbering semantics for formula {tex[:120]!r}; use explicit math delimiters or equation/equation* in source macros")
			descriptor = self.descriptors[markers[0]]
			if descriptor["display"] != display:
				raise FrontendError(f"{descriptor['location']}: macro expansion changed inline/display math semantics")
			tex, labels = _labels(_MARKER.sub("", tex), descriptor["location"])
			_check_math(tex, descriptor["location"])
			if len(labels) > 1:
				raise FrontendError(f"{descriptor['location']}: multiple labels on one equation are unsupported: {labels}")
			if labels and not descriptor["numbered"]:
				raise FrontendError(f"{descriptor['location']}: label {labels[0]!r} is on an unnumbered formula; use equation or remove the label")
			number = len(self.equations) + 1
			identifier = f"eq{number:04d}"
			item = {
				"id": identifier, "kind": "equation", "token": "L2W_MATH_" + identifier,
				"bookmark": f"L2WEq{number:04d}", "tex": tex.strip(), "display": display,
				"numbered": descriptor["numbered"], "label": labels[0] if labels else None,
				"source_location": descriptor["location"],
			}
			if not item["tex"]:
				raise FrontendError(f"{descriptor['location']}: empty formula")
			self.equations.append(item)
			self.register(item)
			return {"t": "Str", "c": item["token"]}
		elif kind == "Link":
			attributes = dict(node["c"][0][2])
			if "reference-type" in attributes:
				return self.reference(attributes["reference-type"], attributes.get("reference", node["c"][2][0].lstrip("#")), in_note)
		elif kind == "Cite":
			raise FrontendError("An unresolved citation remains after citeproc; cannot safely generate the document")
		elif kind in {"RawInline", "RawBlock"}:
			format_name, raw = node["c"]
			if format_name in {"latex", "tex"}:
				reference = re.fullmatch(r"\s*\\(ref|eqref)\s*\{([^{}\\\s]+)\}\s*", raw)
				if reference:
					return self.reference(reference.group(1), reference.group(2), in_note)
				if raw.strip() == r"\maketitle":
					# The title/author/date already exist in Pandoc metadata.
					return {"t": "Null"}
				bibliography_style = re.fullmatch(r"\s*\\bibliographystyle\s*\{([^{}\\]+)\}\s*", raw)
				if bibliography_style:
					self.warnings.append(f"LaTeX bibliography style {bibliography_style.group(1)!r} is replaced by Pandoc's default CSL style")
					return {"t": "Null"}
				if raw.strip() in {r"\centering", r"\raggedright", r"\raggedleft", r"\noindent", r"\indent"}:
					self.warnings.append(f"LaTeX paragraph directive {raw.strip()} is replaced by the configured Word semantic style")
					return {"t": "Null"}
			offset = self.source_text.find(raw)
			location = _location(self.source_text, offset, f"{self.source} (expanded source)") if offset >= 0 else str(self.source)
			message = f"{location}: Unsupported raw {format_name} ({kind}) would be lost by DOCX: {raw[:180]!r}"
			if self.strict:
				raise FrontendError(message + "; rewrite it using supported LaTeX, or use explicit lossy mode to retain it as literal code")
			self.warnings.append(message + "; retained as literal code under lossy mode")
			return {"t": "Code" if kind == "RawInline" else "CodeBlock", "c": [["", ["latex"], []], raw]}
		for key, child in list(node.items()):
			if key != "t":
				node[key] = self.walk(child, in_note=in_note)
		if kind in {"Para", "Plain"} and not node["c"]:
			return {"t": "Null"}
		return node

	def resolve_references(self) -> None:
		for reference in self.references:
			target = self.labels.get(reference["target"])
			if not target:
				raise FrontendError(f"Unresolved \\{reference['command']} target {reference['target']!r}; no corresponding equation/figure/table/heading label")
			if not target.get("numbered", False):
				raise FrontendError(f"Reference {reference['target']!r} points to an unnumbered target")
			if reference["command"] == "eqref" and target["kind"] != "equation":
				raise FrontendError(f"\\eqref{{{reference['target']}}} does not point to an equation")
			reference["kind"] = target["kind"]


def _split_display_paragraphs(value: Any, display_tokens: set[str]) -> Any:
	if isinstance(value, dict):
		return {key: _split_display_paragraphs(child, display_tokens) for key, child in value.items()}
	if not isinstance(value, list):
		return value
	output: list[Any] = []
	for original in value:
		node = _split_display_paragraphs(original, display_tokens)
		if isinstance(node, dict) and node.get("t") in {"Para", "Plain"} and any(item.get("t") == "Str" and item.get("c") in display_tokens for item in node["c"]):
			current: list[dict[str, Any]] = []

			def flush() -> None:
				while current and current[0].get("t") in {"Space", "SoftBreak"}:
					current.pop(0)
				while current and current[-1].get("t") in {"Space", "SoftBreak"}:
					current.pop()
				if current:
					output.append({"t": "Para", "c": current.copy()})
				current.clear()

			for item in node["c"]:
				if item.get("t") == "Str" and item.get("c") in display_tokens:
					flush()
					output.append({"t": "Para", "c": [item]})
				else:
					current.append(item)
			flush()
		else:
			output.append(node)
	return output


def _markdown(ast: dict[str, Any], manifest: dict[str, Any], executable: str, root: Path) -> tuple[str, list[str]]:
	# Use a separate AST so semantic comments do not leak into the Word output.
	# Disable the citations extension: otherwise Pandoc deliberately omits the
	# already rendered reference-list Div from Markdown, losing source content.
	markdown, messages = _pandoc(executable, ["--from=json", "--to=markdown-citations", "--standalone", "--wrap=none", "--sandbox"], json.dumps(ast, ensure_ascii=False), root)
	for equation in manifest["equations"]:
		tex = " ".join(equation["tex"].splitlines()).strip()
		detail = {key: equation[key] for key in ("id", "label", "display", "numbered")}
		replacement = "$" + tex + "$<!-- l2w-equation: " + json.dumps(detail, ensure_ascii=False) + " -->"
		markdown = markdown.replace(equation["token"], replacement)
	for reference in manifest["references"]:
		markdown = markdown.replace(reference["token"], "[[" + reference["command"] + ":" + reference["target"] + "]]")
	for target in manifest["targets"]:
		detail = {key: target[key] for key in ("id", "kind", "label", "numbered")}
		if "short_caption" in target:
			detail["short_caption"] = target["short_caption"]
		markdown = markdown.replace(target["token"], "<!-- l2w-target: " + json.dumps(detail, ensure_ascii=False) + " -->")
	if _RESERVED.search(markdown):
		raise FrontendError("Internal error: Markdown writer altered a semantic placeholder")
	return markdown, messages


def prepare(source: Path, workdir: Path, pandoc: str | None = None, *, strict: bool = True) -> dict[str, Any]:
	"""Prepare all content before DOCX generation, without executing TeX.

	The caller must save ``markdown`` as UTF-8 before passing ``ast`` to a DOCX
	writer. ``workdir`` is accepted as the build-stage interface; this frontend
	itself does not write artifacts. Input files/assets are confined to the
	main source directory, and Pandoc's reader runs with ``--sandbox``.
	"""
	source = Path(source).resolve()
	if not source.is_file():
		raise FrontendError(f"LaTeX source not found: {source}")
	root = source.parent
	executable = find_pandoc(pandoc)
	warnings: list[str] = []
	sources: list[str] = []
	expanded = _expand_inputs(source, root, sources)
	annotated, descriptors = _annotate_math(expanded, source, warnings)
	output, messages = _pandoc(executable, ["--from=latex+raw_tex", "--to=json", "--sandbox"], annotated, root)
	warnings.extend(messages)
	try:
		ast = json.loads(output)
	except json.JSONDecodeError as error:
		raise FrontendError(f"Pandoc returned invalid JSON: {error}") from error
	if ast.get("pandoc-api-version", [0, 0])[:2] < [1, 23]:
		raise FrontendError("Pandoc 3.x with AST API 1.23 or newer is required")
	ast, bibliographies, citation_keys = _bibliography(ast, root, executable, descriptors, warnings)
	transformer = _Transformer(root, descriptors, strict, warnings, source, expanded)
	ast = transformer.walk(ast)
	transformer.resolve_references()
	ast = _split_display_paragraphs(ast, {item["token"] for item in transformer.equations if item["display"]})
	manifest = {
		"version": 1, "source": str(source), "sources": sources, "assets": transformer.assets,
		"bibliographies": bibliographies, "citation_keys": citation_keys,
		"equations": transformer.equations, "references": transformer.references, "targets": transformer.targets,
	}
	markdown, messages = _markdown(ast, manifest, executable, root)
	warnings.extend(messages)
	if transformer.equations:
		warnings.append("Vector/matrix typography is preserved from source; symbol meaning is not inferred. Review bold symbols and actual MathType conversion/layout.")
	return {"ast": ast, "manifest": manifest, "markdown": markdown, "warnings": warnings, "pandoc": executable}
