# author: Codex
"""Safety boundaries independent of Word installation; native acceptance is live."""
import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
import zlib

import pytest

# Source checkout: allow running this file directly without installing the package.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from latex2word import mathtype


def document(path: Path, text: str = "source") -> None:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", f"<document>{text}</document>")


def manifest():
    return {
        "equations": [{"id": "eq0001", "bookmark": "L2WEq0001", "tex": r"x^2",
                       "display": True, "numbered": True, "label": "eq:a"}],
        "references": [{"bookmark": "L2WRef0001", "target": "eq:a", "kind": "equation"}],
    }


@pytest.mark.parametrize("change", ["duplicate", "missing_target", "delimiters", "numbered_inline"])
def test_manifest_validation_precedes_worker(tmp_path, monkeypatch, change):
    source = tmp_path / "input.docx"
    document(source)
    data = manifest()
    if change == "duplicate":
        data["equations"].append(dict(data["equations"][0]))
    elif change == "missing_target":
        data["references"][0]["target"] = "eq:missing"
    elif change == "delimiters":
        data["equations"][0]["tex"] = "$x$"
    else:
        data["equations"][0]["display"] = False
    monkeypatch.setattr(mathtype, "_invoke", lambda *_: pytest.fail("Worker must not run"))
    with pytest.raises(ValueError):
        mathtype.process_document(source, tmp_path / "output.docx", data)


def test_fatal_worker_preserves_existing_output(tmp_path, monkeypatch):
    source, output = tmp_path / "input.docx", tmp_path / "output.docx"
    document(source); document(output, "existing")
    original = output.read_bytes()

    def worker(parameters, timeout):
        Path(parameters["ResultPath"]).write_text(json.dumps({"fatal_error": "probe failed"}))
        document(Path(parameters["OutputPath"]), "partial")
        return subprocess.CompletedProcess([], 1, "", "")

    monkeypatch.setattr(mathtype, "_invoke", worker)
    with pytest.raises(mathtype.MathTypeError, match="probe failed"):
        mathtype.process_document(source, output, manifest())
    assert output.read_bytes() == original
    assert not list(tmp_path.glob(".latex2word-mt-*"))


def test_timeout_cleans_only_recorded_worker(tmp_path, monkeypatch):
    source, output = tmp_path / "input.docx", tmp_path / "output.docx"
    document(source); document(output, "existing")
    original = output.read_bytes()
    calls = []

    def worker(parameters, timeout):
        calls.append(parameters)
        if parameters["Mode"] == "Cleanup":
            assert Path(parameters["ProcessPath"]).read_text() == "owned-pid-record"
            return subprocess.CompletedProcess([], 0, "owned process exited", "")
        Path(parameters["ProcessPath"]).write_text("owned-pid-record")
        raise subprocess.TimeoutExpired("powershell", timeout)

    monkeypatch.setattr(mathtype, "_invoke", worker)
    with pytest.raises(mathtype.MathTypeError, match="timed out"):
        mathtype.process_document(source, output, manifest(), timeout=1)
    assert [c["Mode"] for c in calls] == ["Process", "Cleanup"]
    assert output.read_bytes() == original


def test_success_promotes_verified_staged_copy(tmp_path, monkeypatch):
    source, output = tmp_path / "input.docx", tmp_path / "output.docx"
    document(source)
    original = source.read_bytes()

    def worker(parameters, timeout):
        assert Path(parameters["InputPath"]) != source
        assert Path(parameters["InputPath"]).read_bytes() == original
        document(Path(parameters["OutputPath"]), "converted")
        Path(parameters["ResultPath"]).write_text(json.dumps({"conversion": "converted", "pdf_exported": False}))
        return subprocess.CompletedProcess([], 0, "", "")

    monkeypatch.setattr(mathtype, "_invoke", worker)
    report = mathtype.process_document(source, output, manifest())
    assert report["conversion"] == "converted"
    assert source.read_bytes() == original
    with zipfile.ZipFile(output) as archive:
        assert b"converted" in archive.read("word/document.xml")


def test_bad_worker_package_is_not_promoted(tmp_path, monkeypatch):
    source, output = tmp_path / "input.docx", tmp_path / "output.docx"
    document(source); document(output, "existing")
    original = output.read_bytes()

    def worker(parameters, timeout):
        Path(parameters["OutputPath"]).write_text("not a zip")
        Path(parameters["ResultPath"]).write_text(json.dumps({"conversion": "converted"}))
        return subprocess.CompletedProcess([], 0, "", "")

    monkeypatch.setattr(mathtype, "_invoke", worker)
    with pytest.raises(mathtype.MathTypeError, match="readable Word"):
        mathtype.process_document(source, output, manifest())
    assert output.read_bytes() == original


def test_command_uses_data_not_interpolated_shell_paths():
    command = mathtype._command("powershell.exe", {"InputPath": "a'$(whoami)`b.docx", "Mode": "Process"})
    assert command[:4] == ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive"]
    assert "-ExecutionPolicy" not in command
    assert "whoami" not in " ".join(command)


def test_locked_docx_restores_previous_pdf(tmp_path, monkeypatch):
    source, output = tmp_path / "input.docx", tmp_path / "output.docx"
    pdf = output.with_suffix(".pdf")
    document(source); document(output, "existing")
    pdf.write_bytes(b"previous pdf")
    original = output.read_bytes()

    def worker(parameters, timeout):
        staged = Path(parameters["OutputPath"])
        document(staged, "converted")
        staged.with_suffix(".pdf").write_bytes(b"new pdf")
        Path(parameters["ResultPath"]).write_text(json.dumps({"conversion": "converted", "pdf_exported": True}))
        return subprocess.CompletedProcess([], 0, "", "")

    replace = mathtype.os.replace

    def locked_replace(src, dst):
        if Path(dst) == output:
            raise PermissionError("Word has the destination open")
        replace(src, dst)

    monkeypatch.setattr(mathtype, "_invoke", worker)
    monkeypatch.setattr(mathtype.os, "replace", locked_replace)
    with pytest.raises(PermissionError):
        mathtype.process_document(source, output, manifest(), export_pdf=True)
    assert output.read_bytes() == original
    assert pdf.read_bytes() == b"previous pdf"


@pytest.mark.skipif(os.name != "nt", reason="PDF preview rasterization uses Windows GDI+")
def test_pdf_copy_preserves_dimensions_baseline_and_native_source(tmp_path, monkeypatch):
    """The PDF workaround must never replace the editable primary equation."""
    source, rendered = tmp_path / "native.docx", tmp_path / "rendered.docx"
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    v = "urn:schemas-microsoft-com:vml"
    o = "urn:schemas-microsoft-com:office:office"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xml = f'''<w:document xmlns:w="{w}" xmlns:v="{v}" xmlns:o="{o}" xmlns:r="{r}">
      <w:body><w:p><w:bookmarkStart w:id="1" w:name="L2WEq0001"/>
      <w:r><w:rPr><w:position w:val="-6"/></w:rPr><w:object>
      <v:shapetype id="_x0000_t75"/><v:shape id="eq1" type="#_x0000_t75" style="width:20pt;height:10pt" o:ole="">
      <v:imagedata r:id="rId1"/></v:shape><o:OLEObject ProgID="Equation.DSMT4" r:id="rId2"/>
      </w:object></w:r><w:bookmarkEnd w:id="1"/></w:p>
      <w:p><w:fldSimple w:instr="MACROBUTTON MTPlaceRef preserved-number"/></w:p></w:body></w:document>'''
    rels = f'''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="{r}/image" Target="media/equation.png"/>
      <Relationship Id="rId2" Type="{r}/oleObject" Target="embeddings/equation.bin"/>
      </Relationships>'''
    types = '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="png" ContentType="image/png"/></Types>'''

    def chunk(name, data):
        return struct.pack(">I", len(data)) + name + data + struct.pack(">I", zlib.crc32(name + data))

    pixel = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 6, 0, 0, 0))
             + chunk(b"IDAT", zlib.compress(b"\0\0\0\0\0")) + chunk(b"IEND", b""))
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("[Content_Types].xml", types)
        archive.writestr("word/document.xml", xml)
        archive.writestr("word/_rels/document.xml.rels", rels)
        archive.writestr("word/media/equation.png", pixel)
        archive.writestr("word/embeddings/equation.bin", b"test-only-ole-placeholder")
    original = source.read_bytes()
    worker_script = mathtype._SCRIPT
    harness = tmp_path / "preview-test.ps1"
    harness.write_text('''param([string]$WorkerScript,[string]$Source,[string]$Destination)
Set-StrictMode -Version Latest
$tokens=$null; $errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($WorkerScript,[ref]$tokens,[ref]$errors)
foreach($fn in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Read-ZipXml','Write-ZipXml','New-PdfRenderingCopy')},$true)) {
    Invoke-Expression $fn.Extent.Text
}
New-PdfRenderingCopy $Source $Destination
''', encoding="utf-8")
    monkeypatch.setattr(mathtype, "_SCRIPT", harness)
    result = mathtype._invoke({"WorkerScript": str(worker_script), "Source": str(source), "Destination": str(rendered)}, 20)
    assert result.returncode == 0, result.stderr
    assert result.stdout.strip() == "1"
    assert source.read_bytes() == original
    with zipfile.ZipFile(rendered) as archive:
        document_xml = ET.fromstring(archive.read("word/document.xml"))
        assert not document_xml.findall(f".//{{{o}}}OLEObject")
        shape = document_xml.find(f".//{{{w}}}pict/{{{v}}}shape")
        assert shape is not None and shape.get("style") == "width:20pt;height:10pt"
        assert document_xml.find(f".//{{{w}}}position").get(f"{{{w}}}val") == "-6"
        assert document_xml.find(f".//{{{v}}}shapetype") is not None
        assert document_xml.find(f".//{{{w}}}bookmarkStart").get(f"{{{w}}}name") == "L2WEq0001"
        assert document_xml.find(f".//{{{w}}}fldSimple").get(f"{{{w}}}instr").endswith("preserved-number")
        relationships = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        image_target = next(e.get("Target") for e in relationships if e.get("Id") == "rId1")
        png = archive.read("word/" + image_target)
        assert struct.unpack(">II", png[16:24]) == (167, 84)
