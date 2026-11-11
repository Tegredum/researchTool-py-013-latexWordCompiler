# author: Codex
"""Validated, portable style configuration for the LaTeX-to-Word workflow.

The private ``_explicit`` member returned by :func:`load_config` records which
properties the user supplied. A reference document's existing style and page
properties must only be changed by these explicit overrides. The complete
defaults remain available for creating styles absent from that document.
"""

from __future__ import annotations

import copy
import json
import math
from pathlib import Path
from typing import Any


class ConfigError(ValueError):
	"""The configuration cannot be safely interpreted."""


def _style(name: str, **overrides: Any) -> dict[str, Any]:
	result = {
		"name": name,
		"latin_font": "Times New Roman",
		"east_asia_font": "宋体",
		"size_pt": 12,
		"bold": False,
		"alignment": "justify",
		"space_before_pt": 0,
		"space_after_pt": 0,
		"line_spacing": 1.5,
		"first_line_indent_pt": 0,
		"keep_with_next": False,
	}
	result.update(overrides)
	return result


DEFAULT_CONFIG: dict[str, Any] = {
	"version": 1,
	"reference_docx": None,
	"page": {
		"width_cm": 21,
		"height_cm": 29.7,
		"margin_top_cm": 2.54,
		"margin_bottom_cm": 2.54,
		"margin_left_cm": 2.54,
		"margin_right_cm": 2.54,
	},
	"styles": {
		"body": _style("Normal", first_line_indent_pt=24),
		"title": _style("Title", size_pt=18, bold=True, alignment="center",
						space_after_pt=12, keep_with_next=True),
		"heading1": _style("Heading 1", size_pt=16, bold=True, alignment="left",
						space_before_pt=12, space_after_pt=6, keep_with_next=True),
		"heading2": _style("Heading 2", size_pt=14, bold=True, alignment="left",
						space_before_pt=10, space_after_pt=6, keep_with_next=True),
		"heading3": _style("Heading 3", bold=True, alignment="left",
						space_before_pt=8, space_after_pt=4, keep_with_next=True),
		"caption": _style("Caption", size_pt=10.5, alignment="center", line_spacing=1),
		"figure_caption": _style("Figure Caption", size_pt=10.5, alignment="center",
								line_spacing=1, space_after_pt=6),
		"table_caption": _style("Table Caption", size_pt=10.5, alignment="center",
								line_spacing=1, space_before_pt=6, keep_with_next=True),
		"table": _style("Table Text", size_pt=10.5, alignment="left", line_spacing=1),
		"figure": _style("Figure", alignment="center", line_spacing=1,
						keep_with_next=True),
		"equation": _style("Equation", alignment="left", line_spacing=1,
						space_before_pt=6, space_after_pt=6),
		"code": _style("Source Code", latin_font="Consolas", size_pt=10,
					alignment="left", line_spacing=1),
	},
	"headings": {"numbered": True},
	"captions": {"figure_label": "图", "table_label": "表"},
	"math": {"timeout_seconds": 180},
}


def _object(value: Any, path: str, allowed: set[str]) -> dict[str, Any]:
	if not isinstance(value, dict):
		raise ConfigError(f"{path} must be a JSON object")
	unknown = set(value) - allowed
	if unknown:
		raise ConfigError(f"{path}: unknown key(s): {', '.join(sorted(unknown))}")
	return value


def _string(value: Any, path: str, max_length: int = 128) -> None:
	if not isinstance(value, str) or not value.strip():
		raise ConfigError(f"{path} must be a nonempty string")
	if len(value) > max_length or any(ord(char) < 32 for char in value):
		raise ConfigError(f"{path} must contain at most {max_length} characters and no controls")


def _number(value: Any, path: str, lower: float, upper: float) -> None:
	if isinstance(value, bool) or not isinstance(value, (int, float)):
		raise ConfigError(f"{path} must be a number")
	# Compare first: converting an arbitrarily large JSON integer to float in
	# math.isfinite can raise OverflowError rather than a useful ConfigError.
	if not lower <= value <= upper or not math.isfinite(value):
		raise ConfigError(f"{path} must be finite and between {lower:g} and {upper:g}")


def _boolean(value: Any, path: str) -> None:
	if not isinstance(value, bool):
		raise ConfigError(f"{path} must be true or false")


def _unique_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
	result: dict[str, Any] = {}
	for key, value in pairs:
		if key in result:
			raise ConfigError(f"duplicate JSON key: {key}")
		result[key] = value
	return result


def load_config(path: Path | None = None) -> dict[str, Any]:
	"""Load a UTF-8 JSON file, strictly validate it, and overlay defaults.

	``reference_docx`` resolves relative to the configuration file, never the
	current working directory. ``_explicit`` is generated metadata, not an
	accepted configuration key. There are no shell, registry, or network calls.
	"""
	supplied: dict[str, Any] = {}
	config_path = Path(path).resolve() if path is not None else None
	if config_path is not None:
		try:
			supplied = json.loads(config_path.read_text(encoding="utf-8-sig"),
								object_pairs_hook=_unique_object)
		except (OSError, UnicodeError, json.JSONDecodeError) as exc:
			raise ConfigError(f"Cannot read configuration {config_path}: {exc}") from exc
	_object(supplied, "config", set(DEFAULT_CONFIG))
	version = supplied.get("version", 1)
	if type(version) is not int or version != 1:
		raise ConfigError("version must be the integer 1")

	result = copy.deepcopy(DEFAULT_CONFIG)
	result["_explicit"] = {"page": [], "styles": {}}
	for section in ("page", "headings", "captions", "math"):
		incoming = _object(supplied.get(section, {}), section, set(DEFAULT_CONFIG[section]))
		result[section].update(incoming)
	result["_explicit"]["page"] = sorted(supplied.get("page", {}))
	incoming_styles = _object(supplied.get("styles", {}), "styles", set(DEFAULT_CONFIG["styles"]))
	for semantic, values in incoming_styles.items():
		values = _object(values, f"styles.{semantic}", set(DEFAULT_CONFIG["styles"][semantic]))
		result["styles"][semantic].update(values)
		result["_explicit"]["styles"][semantic] = sorted(values)

	for key, value in result["page"].items():
		_number(value, f"page.{key}", 5 if key in {"width_cm", "height_cm"} else 0,
				100 if key in {"width_cm", "height_cm"} else 50)
	page = result["page"]
	if page["margin_left_cm"] + page["margin_right_cm"] >= page["width_cm"] - 1:
		raise ConfigError("page left/right margins must leave more than 1 cm of usable width")
	if page["margin_top_cm"] + page["margin_bottom_cm"] >= page["height_cm"] - 1:
		raise ConfigError("page top/bottom margins must leave more than 1 cm of usable height")

	for semantic, values in result["styles"].items():
		prefix = f"styles.{semantic}"
		for key in ("name", "latin_font", "east_asia_font"):
			_string(values[key], f"{prefix}.{key}")
		for key in ("bold", "keep_with_next"):
			_boolean(values[key], f"{prefix}.{key}")
		_number(values["size_pt"], f"{prefix}.size_pt", 1, 400)
		for key in ("space_before_pt", "space_after_pt"):
			_number(values[key], f"{prefix}.{key}", 0, 1584)
		_number(values["line_spacing"], f"{prefix}.line_spacing", 0.5, 5)
		_number(values["first_line_indent_pt"], f"{prefix}.first_line_indent_pt", -720, 720)
		if not isinstance(values["alignment"], str) or values["alignment"] not in {
			"left", "center", "right", "justify", "distribute"
		}:
			raise ConfigError(f"{prefix}.alignment must be left, center, right, justify, or distribute")
	_boolean(result["headings"]["numbered"], "headings.numbered")
	for key, value in result["captions"].items():
		_string(value, f"captions.{key}", max_length=32)
	timeout = result["math"]["timeout_seconds"]
	if type(timeout) is not int or not 1 <= timeout <= 3600:
		raise ConfigError("math.timeout_seconds must be an integer between 1 and 3600")

	reference = supplied.get("reference_docx")
	if reference is not None:
		_string(reference, "reference_docx", max_length=32767)
		reference_path = Path(reference).expanduser()
		if not reference_path.is_absolute():
			reference_path = (config_path.parent if config_path is not None else Path.cwd()) / reference_path
		reference_path = reference_path.resolve()
		if reference_path.suffix.lower() not in {".docx", ".dotx"}:
			raise ConfigError("reference_docx must be a .docx or .dotx file")
		if not reference_path.is_file():
			raise ConfigError(f"reference_docx does not exist: {reference_path}")
		result["reference_docx"] = str(reference_path)
	return result
