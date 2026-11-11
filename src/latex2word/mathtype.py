# author: Codex
"""Isolated Windows Word/MathType automation with an atomic output boundary.

The PowerShell worker uses the installed vendor macros.  It never manufactures
OLE data, changes Office security, or writes licensing values.  A successful
native numbering pass is retained if formula conversion subsequently fails.
"""

from __future__ import annotations

import base64
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from typing import Any
import zipfile


_SCRIPT = Path(__file__).with_name("scripts") / "word_mathtype.ps1"
_BOOKMARK = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,39}$")


class MathTypeError(RuntimeError):
	"""The worker could not safely finish a document transaction."""


def _powershell() -> str | None:
	if os.name != "nt":
		return None
	# Windows PowerShell's COM support is tested; use the native 64-bit host.
	system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
	for candidate in (
		system_root / "Sysnative/WindowsPowerShell/v1.0/powershell.exe",
		system_root / "System32/WindowsPowerShell/v1.0/powershell.exe",
	):
		if candidate.is_file():
			return str(candidate)
	return shutil.which("powershell.exe")


def _command(executable: str, parameters: dict[str, Any]) -> list[str]:
	"""Pass paths/data without shell interpolation or changing execution policy.

	Loading our checked-in worker as a ScriptBlock allows normal command-line
	execution under Windows' default Restricted script-file policy.  Group
	policy, constrained language, and Office macro policy remain authoritative.
	"""
	payload = base64.b64encode(
		json.dumps({"script": str(_SCRIPT), "parameters": parameters}).encode("utf-8")
	).decode("ascii")
	bootstrap = (
		"$ErrorActionPreference='Stop';"
		f"$job=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{payload}'))|ConvertFrom-Json;"
		"$p=@{}; $job.parameters.PSObject.Properties|ForEach-Object{$p[$_.Name]=$_.Value};"
		"& ([scriptblock]::Create([IO.File]::ReadAllText($job.script,[Text.Encoding]::UTF8))) @p"
	)
	encoded = base64.b64encode(bootstrap.encode("utf-16-le")).decode("ascii")
	return [executable, "-NoLogo", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded]


def _invoke(parameters: dict[str, Any], timeout: int) -> subprocess.CompletedProcess[str]:
	executable = _powershell()
	if executable is None:
		raise MathTypeError("Word/MathType automation requires Windows and Windows PowerShell.")
	return subprocess.run(
		_command(executable, parameters),
		capture_output=True,
		text=True,
		encoding="utf-8",
		errors="replace",
		timeout=timeout,
		check=False,
		creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
	)


def _validate_manifest(manifest: dict[str, Any]) -> None:
	equations = manifest.get("equations", [])
	references = manifest.get("references", [])
	if not isinstance(equations, list) or not isinstance(references, list):
		raise ValueError("Manifest equations and references must be lists.")
	ids: set[str] = set()
	bookmarks: set[str] = set()
	labels: set[str] = set()
	for equation in equations:
		eq_id = equation.get("id", "")
		if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", eq_id) or eq_id in ids:
			raise ValueError(f"Invalid or duplicate equation id: {eq_id!r}")
		ids.add(eq_id)
		tex = equation.get("tex", "")
		if not isinstance(tex, str) or not tex.strip() or any(c in tex for c in "\r\n$"):
			raise ValueError(f"Equation {eq_id}: tex must be an undelimited, single-line TeX body.")
		if re.search(r"\\(?:tag\*?\s*\{|begin\s*\{align\*?\})", tex):
			raise ValueError(f"Equation {eq_id}: tag/align must be normalized before MathType.")
		if equation.get("numbered") and not equation.get("display"):
			raise ValueError(f"Equation {eq_id}: only display formulas can be numbered.")
		label = equation.get("label")
		if label:
			if label in labels:
				raise ValueError(f"Duplicate equation label: {label!r}")
			labels.add(label)
	for record in [*equations, *references]:
		bookmark = record.get("bookmark", "")
		if not _BOOKMARK.fullmatch(bookmark) or bookmark in bookmarks:
			raise ValueError(f"Invalid or duplicate bookmark: {bookmark!r}")
		bookmarks.add(bookmark)
	numbered_labels = {e.get("label") for e in equations if e.get("numbered") and e.get("label")}
	for reference in references:
		if reference.get("kind") == "equation" and reference.get("target") not in numbered_labels:
			raise ValueError(f"Reference has no numbered equation target: {reference.get('target')!r}")


def doctor(*, timeout: int = 20) -> dict[str, Any]:
	"""Read registration/context metadata only; do not launch Word or change state."""
	if _powershell() is None:
		return {"available": False, "platform": os.name, "reason": "Windows PowerShell is unavailable", "functional_test": "not_run"}
	try:
		result = _invoke({"Mode": "Doctor"}, timeout)
		report = json.loads(result.stdout.lstrip("\ufeff").strip())
		if not isinstance(report, dict):
			raise ValueError("Diagnostic response is not an object")
		return report
	except (OSError, subprocess.TimeoutExpired, ValueError, MathTypeError) as exc:
		return {"available": False, "reason": str(exc), "functional_test": "not_run"}


def _check_docx(path: Path) -> None:
	try:
		with zipfile.ZipFile(path) as archive:
			if "word/document.xml" not in archive.namelist() or archive.testzip() is not None:
				raise MathTypeError("Worker output is not an intact Word package.")
	except (OSError, zipfile.BadZipFile) as exc:
		raise MathTypeError(f"Worker output is not a readable Word document: {exc}") from exc


def process_document(
	input_docx: Path,
	output_docx: Path,
	manifest: dict[str, Any],
	*,
	timeout: int = 180,
	export_pdf: bool = False,
) -> dict[str, Any]:
	"""Process a copy, then atomically replace ``output_docx`` after verification.

	``tex`` contains the body without delimiters.  Equation bookmarks must cover
	exactly ``$tex$``; numbered paragraphs must contain ``L2W_NUMBER_<id>``.
	A returned ``tex_retained`` or ``unavailable`` status is explicit fallback,
	never proof of complete MathType compatibility.  Timeouts/fatal errors leave
	both the input and any existing output intact and raise ``MathTypeError``.
	"""
	source = Path(input_docx).resolve(strict=True)
	destination = Path(output_docx).resolve()
	if source.suffix.lower() != ".docx" or destination.suffix.lower() != ".docx":
		raise ValueError("MathType input and output must use .docx.")
	if timeout <= 0:
		raise ValueError("timeout must be positive")
	_validate_manifest(manifest)
	_check_docx(source)
	if not destination.parent.is_dir():
		raise FileNotFoundError(destination.parent)
	# Office can briefly retain file handles while its process exits. Never let
	# temporary-directory cleanup hide the primary, actionable worker failure.
	with tempfile.TemporaryDirectory(prefix=".latex2word-mt-", dir=destination.parent, ignore_cleanup_errors=True) as tmp:
		stage = Path(tmp)
		copied_input = stage / "input.docx"
		staged_output = stage / "output.docx"
		result_path = stage / "result.json"
		owner_path = stage / "word-process.json"
		manifest_path = stage / "manifest.json"
		shutil.copy2(source, copied_input)
		normalized_manifest = {**manifest, "equations": manifest.get("equations", []), "references": manifest.get("references", [])}
		manifest_path.write_text(json.dumps(normalized_manifest, ensure_ascii=False), encoding="utf-8")
		parameters = {
			"Mode": "Process",
			"InputPath": str(copied_input),
			"OutputPath": str(staged_output),
			"ManifestPath": str(manifest_path),
			"ResultPath": str(result_path),
			"ProcessPath": str(owner_path),
			"ExportPdf": export_pdf,
		}
		try:
			result = _invoke(parameters, timeout)
		except subprocess.TimeoutExpired as exc:
			cleanup = "No owned Word process was recorded."
			if owner_path.exists():
				try:
					cleanup_result = _invoke({"Mode": "Cleanup", "ProcessPath": str(owner_path)}, 20)
					cleanup = cleanup_result.stdout.strip() or cleanup_result.stderr.strip()
				except (OSError, subprocess.TimeoutExpired, MathTypeError) as cleanup_exc:
					cleanup = f"Owned-process cleanup could not finish: {cleanup_exc}"
			raise MathTypeError(f"MathType timed out after {timeout}s; existing output preserved. {cleanup} Staging path if still locked: {stage}") from None
		except OSError as exc:
			raise MathTypeError(f"Cannot launch MathType worker: {exc}") from exc
		if not result_path.exists():
			raise MathTypeError(f"MathType worker did not return a report: {result.stderr.strip() or result.stdout.strip()}")
		try:
			report = json.loads(result_path.read_text(encoding="utf-8-sig"))
		except (OSError, ValueError) as exc:
			raise MathTypeError("MathType worker returned an invalid report.") from exc
		if result.returncode != 0 or report.get("fatal_error"):
			raise MathTypeError(report.get("fatal_error") or result.stderr.strip() or "MathType worker failed.")
		_check_docx(staged_output)
		staged_pdf = staged_output.with_suffix(".pdf")
		if report.get("pdf_exported") and not staged_pdf.is_file():
			raise MathTypeError("Worker reported a PDF but did not create it.")
		# Each artifact is promoted only after the Word transaction is complete.
		# If replacing a locked DOCX fails, restore the previous PDF as well.
		pdf_path = destination.with_suffix(".pdf")
		pdf_backup = stage / "previous.pdf"
		pdf_promoted = False
		if report.get("pdf_exported"):
			if pdf_path.exists():
				shutil.copy2(pdf_path, pdf_backup)
			os.replace(staged_pdf, pdf_path)
			pdf_promoted = True
			report["pdf"] = str(pdf_path)
		try:
			os.replace(staged_output, destination)
		except OSError:
			if pdf_promoted:
				if pdf_backup.exists():
					os.replace(pdf_backup, pdf_path)
				else:
					pdf_path.unlink()
			raise
		report["output"] = str(destination)
		return report
