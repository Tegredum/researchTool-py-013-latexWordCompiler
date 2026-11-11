# encoding: utf-8
# author: Codex, DSH
"""Integration checks against a real Pandoc; no Word/MathType required."""

from __future__ import annotations

import base64
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest

# Source checkout: allow running this file directly without installing the package.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from latex2word.frontend import FrontendError, prepare

PANDOC = os.environ.get("PANDOC") or shutil.which("pandoc")
PNG_FIXTURE = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAICAIAAAB/FOjAAAAAE0lEQVR4nGNQSFhAEmIY1UALDQAH+pABFcYqegAAAABJRU5ErkJggg=="

@unittest.skipUnless(PANDOC, "Pandoc is required for frontend integration tests")
class FrontendTests(unittest.TestCase):
	def setUp(self) -> None:
		self.temporary = tempfile.TemporaryDirectory()
		self.root = Path(self.temporary.name)
		self.addCleanup(self.temporary.cleanup)

	def convert(self, body: str, *, preamble: str = "", strict: bool = True):
		source = self.root / "main.tex"
		source.write_text("\\documentclass{article}\n" + preamble + "\n\\begin{document}\n" + body + "\n\\end{document}\n", encoding="utf-8")
		return prepare(source, self.root / "build", PANDOC, strict=strict)

	def test_math_macros_numbering_and_forward_references(self):
		result = self.convert(
			r"""\section{介绍}\label{sec:intro}
中文Word不加空格。See \eqref{eq:model} and \ref{sec:intro}.
Inline $\vect{x}$.
\begin{equation}\dot{\vect{x}}=\vect{A}\vect{x}\label{eq:model}\end{equation}
\[z=2\]
\begin{equation*}q=3\end{equation*}
""", preamble=r"\newcommand{\vect}[1]{\mathbf{#1}}")
		equations = result["manifest"]["equations"]
		self.assertEqual([item["numbered"] for item in equations], [False, True, False, False])
		self.assertEqual([item["display"] for item in equations], [False, True, True, True])
		self.assertEqual(equations[1]["label"], "eq:model")
		self.assertIn(r"\mathbf{A}", equations[1]["tex"])
		self.assertNotIn(r"\label", equations[1]["tex"])
		self.assertNotIn('"Math"', json.dumps(result["ast"]))
		self.assertEqual([ref["kind"] for ref in result["manifest"]["references"]], ["equation", "heading"])
		self.assertIn("中文Word不加空格", result["markdown"])
		self.assertNotIn("$$", result["markdown"])
		self.assertNotIn("L2W_MATH", result["markdown"])
		self.assertIn("[[eqref:eq:model]]", result["markdown"])
		display = {item["token"] for item in equations if item["display"]}
		blocks = result["ast"]["blocks"]
		self.assertEqual(sum(block.get("c") == [{"t": "Str", "c": token}] for block in blocks for token in display), 3)

	def test_local_input_macros_and_verbatim_are_preserved(self):
		(self.root / "macros.tex").write_text(r"\newcommand{\twice}[1]{2#1}", encoding="utf-8")
		result = self.convert(r"""$\twice{x}$
\begin{verbatim}
$literal$ \input{not-a-file} \write18{not-executed}
\end{verbatim}
""", preamble=r"\input{macros}")
		self.assertEqual(result["manifest"]["equations"][0]["tex"], "2x")
		self.assertEqual(len(result["manifest"]["sources"]), 2)
		self.assertIn("$literal$", result["markdown"])

	def test_macro_containing_whole_formula(self):
		result = self.convert(r"\mymath and \mymath", preamble=r"\newcommand{\mymath}{\(\mathbf{x}\)}")
		self.assertEqual(len(result["manifest"]["equations"]), 2)
		self.assertNotEqual(*[item["token"] for item in result["manifest"]["equations"]])
		result = self.convert(r"Text \ensuremath{\mathbf{x}} and literal \$20 and \verb|$code$|.")
		self.assertEqual(len(result["manifest"]["equations"]), 1)
		self.assertIn("$code$", result["markdown"])

	def test_comments_do_not_create_tags_labels_or_math(self):
		result = self.convert("% $not math$\n" + r"\begin{equation}x=1 % \tag{A} \label{eq:fake}" + "\n" + r"\label{eq:real}\end{equation}")
		self.assertEqual(len(result["manifest"]["equations"]), 1)
		self.assertEqual(result["manifest"]["equations"][0]["label"], "eq:real")

	def test_independent_align_rows_and_continuation_rejection(self):
		result = self.convert(r"\begin{align} a &= b \label{eq:a} \\ c &= d \label{eq:c}\end{align}")
		self.assertEqual([item["label"] for item in result["manifest"]["equations"]], ["eq:a", "eq:c"])
		self.assertTrue(all("&" not in item["tex"] for item in result["manifest"]["equations"]))
		self.assertTrue(any("normalized align" in warning for warning in result["warnings"]))
		with self.assertRaisesRegex(FrontendError, "continuation|ambiguous"):
			self.convert(r"\begin{align} a &= b \\ &= c\end{align}")

	def test_caption_targets_and_native_table_and_image_ast(self):
		(self.root / "figure.png").write_bytes(base64.b64decode(PNG_FIXTURE))
		result = self.convert(r"""See \ref{fig:one}, \ref{tbl:one}.
\begin{figure}\includegraphics[width=2cm]{figure.png}\caption{A picture}\label{fig:one}\end{figure}
\begin{table}\caption{A table}\label{tbl:one}\begin{tabular}{ll}a&b\\c&d\end{tabular}\end{table}
""")
		targets = result["manifest"]["targets"]
		self.assertEqual([item["kind"] for item in targets], ["figure", "table"])
		self.assertEqual([item["label"] for item in targets], ["fig:one", "tbl:one"])
		self.assertEqual([item["kind"] for item in result["manifest"]["references"]], ["figure", "table"])
		self.assertIn('"Table"', json.dumps(result["ast"]))
		self.assertIn((self.root / "figure.png").as_posix(), json.dumps(result["ast"]))

	def test_unknown_tex_is_never_silently_dropped(self):
		with self.assertRaisesRegex(FrontendError, "unknowncommand"):
			self.convert(r"Before \unknowncommand{important} after.")
		result = self.convert(r"Before \unknowncommand{important} after.", strict=False)
		self.assertIn("unknowncommand", result["markdown"])
		self.assertTrue(any("literal code" in warning for warning in result["warnings"]))

	def test_rejects_missing_and_duplicate_references(self):
		with self.assertRaisesRegex(FrontendError, "Unresolved"):
			self.convert(r"See \eqref{eq:missing}.")
		with self.assertRaisesRegex(FrontendError, "Duplicate"):
			self.convert(r"\begin{equation}a=b\label{eq:x}\end{equation}\begin{equation}c=d\label{eq:x}\end{equation}")
		with self.assertRaisesRegex(FrontendError, "unnumbered"):
			self.convert(r"\begin{equation*}a=b\label{eq:x}\end{equation*}")

	def test_rejects_tags_nested_layout_and_fourth_heading(self):
		for body in (r"\[x=2\tag{A}\]", r"\begin{minipage}{2cm}text\end{minipage}", r"\[\begin{aligned}a&=b\\c&=d\end{aligned}\]", r"\paragraph{Fourth level} text"):
			with self.subTest(body=body), self.assertRaises(FrontendError):
				self.convert(body)

	def test_security_source_and_asset_boundaries(self):
		for body in (r"\input{https://example.com/a.tex}", r"\input{../outside.tex}", r"\includegraphics{https://example.com/image.png}", r"\write18{anything}"):
			with self.subTest(body=body), self.assertRaises(FrontendError):
				self.convert(body)
		(self.root / "recursive.tex").write_text(r"\input{recursive}", encoding="utf-8")
		with self.assertRaisesRegex(FrontendError, "Recursive"):
			self.convert(r"\input{recursive}")

	def test_reserved_tokens_and_note_math_fail_early(self):
		for token in ("L2W_MATH_eq0001", "L2W_NUMBER_eq0001", "L2W_custom"):
			with self.subTest(token=token), self.assertRaisesRegex(FrontendError, "reserved"):
				self.convert(token)
		with self.assertRaisesRegex(FrontendError, "footnote"):
			self.convert(r"Text\footnote{Formula $x=1$}.")

	def test_bibliography_citeproc_and_new_math_remain_tex(self):
		(self.root / "refs.bib").write_text(r"@article{doe2020,author={Doe, John},title={A result for $\alpha$},journal={Journal},year={2020}}", encoding="utf-8")
		result = self.convert(r"See \citet{doe2020} and \citep{doe2020}.\bibliographystyle{plainnat}\bibliography{refs}")
		self.assertEqual(result["manifest"]["citation_keys"], ["doe2020"])
		self.assertEqual(result["manifest"]["bibliographies"], [str(self.root / "refs.bib")])
		self.assertIn("2020", result["markdown"])
		self.assertIn("Doe", result["markdown"])
		self.assertIn("Journal", result["markdown"])
		self.assertNotIn("@doe2020", result["markdown"])
		self.assertEqual(result["manifest"]["equations"][0]["tex"], r"\alpha")
		self.assertFalse(result["manifest"]["equations"][0]["numbered"])
		self.assertNotIn('"Math"', json.dumps(result["ast"]))
		self.assertTrue(any("default CSL" in warning for warning in result["warnings"]))
		with self.assertRaisesRegex(FrontendError, "Citeproc"):
			self.convert(r"See \cite{missing}.\bibliography{refs}")
		with self.assertRaisesRegex(FrontendError, "local .bib"):
			self.convert(r"See \cite{doe2020}.")
		with self.assertRaisesRegex(FrontendError, "network/URI"):
			self.convert(r"See \cite{doe2020}.", preamble=r"\addbibresource{https://example.com/refs.bib}")

	def test_biblatex_and_short_caption_handling(self):
		(self.root / "refs.bib").write_text("@book{doe2020,author={Doe, John},title={A book},year={2020}}", encoding="utf-8")
		result = self.convert(r"See \autocite{doe2020}.\printbibliography", preamble=r"\addbibresource{refs.bib}")
		self.assertIn("A Book", result["markdown"])
		self.assertNotIn("printbibliography", result["markdown"])
		result = self.convert(r"\begin{table}\caption[Short caption]{Long caption}\begin{tabular}{l}x\end{tabular}\end{table}")
		self.assertIn("Short caption", result["markdown"])
		with self.assertRaisesRegex(FrontendError, "short table caption"):
			self.convert(r"\begin{table}\caption[Short $s$]{Long $l$}\begin{tabular}{l}x\end{tabular}\end{table}")

if __name__ == "__main__":
	unittest.main()
