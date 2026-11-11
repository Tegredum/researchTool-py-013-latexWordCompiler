# author: Codex
# Windows PowerShell 5.1 / installed MathType 6.9 Word add-in.
# Vendor macros remain in their installed templates; none are embedded here.
param(
	[ValidateSet('Doctor','Process','Cleanup')][string]$Mode = 'Doctor',
	[string]$InputPath, [string]$OutputPath, [string]$ManifestPath,
	[string]$ResultPath, [string]$ProcessPath, [bool]$ExportPdf = $false
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

function Write-JsonFile($Path, $Value) {
	[IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
}
function Release-Com($Value) {
	if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
	}
}
function Close-Document($Document) {
	if ($null -ne $Document) { $save = 0; $Document.Close([ref]$save); Release-Com $Document }
}
function Open-Document($Word, [string]$Path) {
	$confirm = $false; $readOnly = $false; $recent = $false
	return $Word.Documents.Open([ref]$Path, [ref]$confirm, [ref]$readOnly, [ref]$recent)
}
function Save-DocumentAs($Document, [string]$Path) {
	$format = 16
	$Document.SaveAs2([ref]$Path, [ref]$format)
}
function Read-RegistryDefault([string]$Subkey) {
	$key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey($Subkey, $false)
	try { if ($key) { return [string]$key.GetValue('') } } finally { if ($key) { $key.Dispose() } }
	return $null
}
function Get-EnvironmentReport {
	$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
	$wordClsid = Read-RegistryDefault 'Word.Application\CLSID'
	$mathClsid = Read-RegistryDefault 'Equation.DSMT4\CLSID'
	$wordServer = if ($wordClsid) { Read-RegistryDefault ('CLSID\' + $wordClsid + '\LocalServer32') } else { $null }
	$mathServer = if ($mathClsid) { Read-RegistryDefault ('CLSID\' + $mathClsid + '\LocalServer32') } else { $null }
	$contextOK = $identity.Name -notmatch '\\CodexSandbox'
	return [ordered]@{
		available = [bool]($wordClsid -and $mathClsid -and $contextOK)
		word_registered = [bool]$wordClsid; mathtype_registered = [bool]$mathClsid
		word_server = $wordServer; mathtype_server = $mathServer
		identity = $identity.Name; desktop_context = $contextOK
		process_64bit = [Environment]::Is64BitProcess
		functional_test = 'not_run'
		note = 'Registration is not proof of usable macros; process_document performs a live probe. No Word instance was opened by doctor.'
	}
}
function Stop-OwnedWord([string]$OwnerPath) {
	if (-not (Test-Path -LiteralPath $OwnerPath)) { return 'No owned Word process recorded.' }
	$owner = [IO.File]::ReadAllText($OwnerPath) | ConvertFrom-Json
	$process = Get-Process -Id $owner.pid -ErrorAction SilentlyContinue
	if (-not $process) { return 'Owned Word process already exited.' }
	if ($process.ProcessName -ne 'WINWORD' -or $process.StartTime.ToUniversalTime().Ticks.ToString() -ne $owner.started_ticks) {
		throw 'PID identity changed; refusing to terminate a different process.'
	}
	Stop-Process -Id $process.Id -ErrorAction Stop
	[void]$process.WaitForExit(10000)
	return ('Stopped only owned Word process ' + $process.Id)
}
if ($Mode -eq 'Doctor') { Get-EnvironmentReport | ConvertTo-Json -Depth 6; return }
if ($Mode -eq 'Cleanup') { Stop-OwnedWord $ProcessPath; return }

function Initialize-MathType($Word) {
	# The supported profile is deliberately narrow. Template names are discovered
	# from this Word instance, not copied from another machine's installation path.
	$templates = @()
	for ($i = 1; $i -le $Word.AddIns.Count; $i++) {
		$addin = $Word.AddIns.Item($i)
		try {
			if ($addin.Installed -and $addin.Name -match '^MathType Commands 6 For Word (2010|2013|2016)\.dotm$') {
				$templates += [string]$addin.Name
			}
		} finally { Release-Com $addin }
	}
	$templates = @($templates | Sort-Object -Unique -Descending)
	if ($templates.Count -eq 0) { throw 'No loaded MathType 6 Word template from the supported profile was found.' }
	$name = $templates[0]
	[void]$Word.Run(("'" + $name + "'!AutoExec.Main"))
	if (-not $Word.Run('IsFullFunctionality')) { throw 'MathType reports full functionality unavailable after AutoExec initialization.' }
	[void]$Word.Run('UpdateState')
	return $name
}
function Set-NumberFormat($Word, $Document) {
	# This is the vendor's documented per-document token format, read by its
	# native insertion command. No fields, bookmark IDs, or registry defaults
	# are reconstructed. Chapter/section numbering is intentionally not inferred.
	# The installed Office PIA throws even when getting CustomDocumentProperties
	# from PowerShell. Use the vendor's public document-property helper, verified
	# in its installed WordCmds.dot source and through a live write/read probe.
	$key = 'MTEquationNumber2'; $value = '(#E1)'
	$result = $Word.Run('MTLib.WriteDocPropString', [ref]$Document, [ref]$key, [ref]$value)
	if ($result -ne 0 -or $Word.Run('MTEqnNumFormat.GetEqnNumFormat', [ref]$Document) -ne $value) {
		throw 'MathType could not store the document equation-number format.'
	}
}
function Get-BookmarkRange($Document, [string]$Name) {
	if (-not $Document.Bookmarks.Exists($Name)) { throw ('Required bookmark is missing: ' + $Name) }
	return $Document.Bookmarks.Item($Name).Range.Duplicate
}
function Find-Text($Range, [string]$Text) {
	$found = $Range.Duplicate
	$find = $found.Find
	$find.ClearFormatting(); $find.Text = $Text; $find.Forward = $true
	$find.Wrap = 0; $find.MatchCase = $true; $find.MatchWildcards = $false
	if (-not $find.Execute()) { throw ('Expected placeholder is missing: ' + $Text) }
	return $found
}
function Get-NativeNumberField($Range) {
	for ($i = 1; $i -le $Range.Fields.Count; $i++) {
		$field = $Range.Fields.Item($i)
		if ($field.Code.Text -match '^\s*MACROBUTTON\s+MTPlaceRef\b') { return $field }
		Release-Com $field
	}
	throw 'MathType did not create its native MACROBUTTON MTPlaceRef number.'
}
function Get-WholeFieldRange($Document, $Field) {
	# MACROBUTTON fields can omit a separator and have an empty Result.  Their
	# Result.End is already after the closing delimiter in this Word version;
	# adding one selects the following character and breaks MTPlaceRef's search.
	# Let Word select its complete field instead of inferring field boundaries.
	$Field.Select()
	return $Document.Application.Selection.Range.Duplicate
}
function Get-NativeDisplayValue($Field) {
	# Outer MACROBUTTON.Result can be empty. Display counters live in its Code.
	$code = $Field.Code
	for ($i = 1; $i -le $code.Fields.Count; $i++) {
		$nested = $code.Fields.Item($i)
		if ($nested.Code.Text -match '^\s*SEQ\s+MTEqn\s+\\c\b') { return $nested.Result.Text.Trim() }
	}
	throw 'Native number lacks its displayed MTEqn counter.'
}
function Insert-NativeNumber($Word, $Document, $Marker, [string]$BookmarkName) {
	$start = $Marker.Start
	$paragraph = $Marker.Paragraphs.Item(1).Range.Duplicate
	$Marker.Text = ''
	$Document.Range($start, $start).Select()
	[void]$Word.Run('MTCommand_InsertEqnNum')
	$field = Get-NativeNumberField $paragraph
	$whole = Get-WholeFieldRange $Document $field
	[void]$Document.Bookmarks.Add($BookmarkName, $whole)
	return $field
}
function Insert-NativeReference($Word, $Document, $Placeholder, $NumberRange) {
	if ($Document.Bookmarks.Exists('MTReference')) { throw 'Unexpected pending MathType reference bookmark.' }
	$position = $Placeholder.Start
	[void]$Document.Bookmarks.Add('MTReference', $Placeholder)
	if (-not $Document.Bookmarks.Exists('MTReference')) { throw 'Could not mark the native reference insertion range.' }
	$before = $Document.Bookmarks.Item('MTReference').Range.Text
	$numberField = Get-NativeNumberField $NumberRange
	$numberField.Select()
	[void]$Word.Run('MTPlaceRef')
	if ($Document.Bookmarks.Exists('MTReference')) { throw ('MathType did not complete reference insertion; placeholder=' + $before + '; number range=' + $NumberRange.Text + '; active=' + $Word.ActiveDocument.Name + '; intended=' + $Document.Name) }
	$paragraph = $Document.Range($position, $position).Paragraphs.Item(1).Range
	$found = $false
	for ($i = 1; $i -le $paragraph.Fields.Count; $i++) {
		$field = $paragraph.Fields.Item($i)
		if ($field.Code.Text -match '^\s*GOTOBUTTON\b' -and $field.Code.Text -match '\bREF\s+') { $found = $true }
	}
	if (-not $found) { throw 'MathType reference lacks its native GOTOBUTTON / REF structure.' }
}
function Update-NativeNumbers($Word, $Document) {
	$Document.Activate()
	$Document.Range(0, 0).Select() # A nonempty selection changes the vendor update scope.
	[void]$Word.Run('MTCommand_UpdateEqns')
}
function Test-NativeNumbering($Word) {
	$sample = $null
	try {
		$sample = $Word.Documents.Add()
		Set-NumberFormat $Word $sample
		$sample.Content.Text = "A NUMBER_A`rB NUMBER_B`rREF_B`r"
		$a = Insert-NativeNumber $Word $sample (Find-Text $sample.Content 'NUMBER_A') 'L2WTestA'
		$b = Insert-NativeNumber $Word $sample (Find-Text $sample.Content 'NUMBER_B') 'L2WTestB'
		Insert-NativeReference $Word $sample (Find-Text $sample.Content 'REF_B') (Get-BookmarkRange $sample 'L2WTestB')
		Update-NativeNumbers $Word $sample
		if ((Get-NativeDisplayValue $a) -ne '1' -or (Get-NativeDisplayValue $b) -ne '2') { throw 'Native sample numbering was not continuous.' }
		# Insert at the middle and at the end using the same vendor command.
		$middle = (Get-BookmarkRange $sample 'L2WTestB').Paragraphs.Item(1).Range.Start
		$sample.Range($middle, $middle).Text = "MID_NUMBER`r"
		$mid = Insert-NativeNumber $Word $sample (Find-Text $sample.Content 'MID_NUMBER') 'L2WTestMid'
		$tail = $sample.Content.End - 1
		$sample.Range($tail, $tail).Text = "END_NUMBER`r"
		$end = Insert-NativeNumber $Word $sample (Find-Text $sample.Content 'END_NUMBER') 'L2WTestEnd'
		Update-NativeNumbers $Word $sample
		$b = Get-NativeNumberField (Get-BookmarkRange $sample 'L2WTestB')
		if ((Get-NativeDisplayValue $b) -ne '3' -or (Get-NativeDisplayValue $end) -ne '4') { throw 'Native insertion failed to renumber the sample.' }
		# Delete a preceding number and verify the referenced target follows.
		(Get-BookmarkRange $sample 'L2WTestMid').Paragraphs.Item(1).Range.Delete() | Out-Null
		Update-NativeNumbers $Word $sample
		$b = Get-NativeNumberField (Get-BookmarkRange $sample 'L2WTestB')
		if ((Get-NativeDisplayValue $b) -ne '2') { throw 'Native deletion failed to renumber the sample.' }
		return [ordered]@{
			native_insert = $true; middle_and_end_insert = $true; delete_and_update = $true
			native_reference_structure = $true
			move_and_reference_tracking = 'manual_check_required'
			format_change = 'manual_check_required'
			scope = 'Disposable preflight sample; final-document acceptance remains separate.'
		}
	} finally { Close-Document $sample }
}
function Update-AllStories($Document) {
	# Header/footer/textbox stories can have linked NextStoryRange chains.
	foreach ($first in $Document.StoryRanges) {
		$story = $first
		while ($null -ne $story) {
			[void]$story.Fields.Update()
			$story = $story.NextStoryRange
		}
	}
}
function Set-TargetView($Document) {
	# These are this document window's display switches, not paragraph/page
	# controls or persistent Office-wide options. Keep pagination constraints.
	$Document.Activate()
	$view = $Document.ActiveWindow.View
	$view.ShowAll = $false
	$view.ShowParagraphs = $false
	$view.ShowFieldCodes = $false
	return (-not $view.ShowAll -and -not $view.ShowParagraphs -and -not $view.ShowFieldCodes)
}
function Read-ZipXml($Archive, [string]$Name) {
	$entry = $Archive.GetEntry($Name)
	if (-not $entry) { throw ('Missing package part: ' + $Name) }
	$reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8)
	try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Write-ZipXml($Archive, [string]$Name, $Xml) {
	$entry = $Archive.GetEntry($Name)
	if ($entry) { $entry.Delete() }
	$writer = [IO.StreamWriter]::new($Archive.CreateEntry($Name).Open(), [Text.UTF8Encoding]::new($false))
	try { $Xml.Save($writer) } finally { $writer.Dispose() }
}
function New-PdfRenderingCopy([string]$Source, [string]$Destination) {
	# Word's Windows PDF exporter can silently omit MT Extra glyphs in MathType
	# WMF records (e.g. the hat accent), although both OLE and preview are correct.
	# Rasterize equation previews ONLY in this disposable PDF rendering copy.
	# The delivered DOCX and its editable Equation.DSMT4 objects stay untouched.
	Add-Type -AssemblyName System.Drawing
	Add-Type -AssemblyName System.IO.Compression
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	[IO.File]::Copy($Source, $Destination, $true)
	$archive = [IO.Compression.ZipFile]::Open($Destination, [IO.Compression.ZipArchiveMode]::Update)
	$count = 0
	$imagePrefix = 'media/l2w-pdf-' + [Guid]::NewGuid().ToString('N') + '-'
	try {
		$parts = @($archive.Entries | Where-Object { $_.FullName -match '^word/(document|header[0-9]+|footer[0-9]+|footnotes|endnotes)\.xml$' } | ForEach-Object { $_.FullName })
		foreach ($part in $parts) {
			$xml = Read-ZipXml $archive $part
			$ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
			$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
			$ns.AddNamespace('o','urn:schemas-microsoft-com:office:office')
			$ns.AddNamespace('v','urn:schemas-microsoft-com:vml')
			$nodes = @($xml.SelectNodes('//w:object[o:OLEObject[@ProgID="Equation.DSMT4"]]', $ns))
			if ($nodes.Count -eq 0) { continue }
			$relsName = 'word/_rels/' + [IO.Path]::GetFileName($part) + '.rels'
			$rels = Read-ZipXml $archive $relsName
			foreach ($node in $nodes) {
				$shape = $node.SelectSingleNode('v:shape', $ns)
				$image = $shape.SelectSingleNode('v:imagedata', $ns)
				$rid = $image.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
				$rel = @($rels.DocumentElement.ChildNodes | Where-Object { $_.Id -eq $rid })[0]
				if (-not $rel -or $rel.GetAttribute('TargetMode') -eq 'External') { throw 'Equation preview is missing or external.' }
				$target = [string]$rel.Target
				if ($target -notmatch '^media/[^/\\]+\.(wmf|emf|png)$') { throw ('Unsupported equation preview target: ' + $target) }
				$widthMatch = [regex]::Match($shape.GetAttribute('style'), '(?:^|;)\s*width:([0-9.]+)pt')
				$heightMatch = [regex]::Match($shape.GetAttribute('style'), '(?:^|;)\s*height:([0-9.]+)pt')
				if (-not $widthMatch.Success -or -not $heightMatch.Success) { throw 'Equation preview has no explicit point dimensions.' }
				$width = [double]::Parse($widthMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
				$height = [double]::Parse($heightMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture)
				$pixelsX = [int][Math]::Ceiling($width * 600 / 72)
				$pixelsY = [int][Math]::Ceiling($height * 600 / 72)
				if ($pixelsX -lt 1 -or $pixelsY -lt 1 -or $pixelsX -gt 10000 -or $pixelsY -gt 10000) { throw 'Equation preview dimensions exceed the PDF rasterization limit.' }
				$sourceStream = $archive.GetEntry('word/' + $target).Open()
				$memory = [IO.MemoryStream]::new()
				$metafile = $null; $bitmap = $null; $graphics = $null; $png = $null
				try {
					$sourceStream.CopyTo($memory); $memory.Position = 0
					$metafile = [Drawing.Image]::FromStream($memory)
					$bitmap = [Drawing.Bitmap]::new($pixelsX, $pixelsY, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
					$bitmap.SetResolution(600, 600)
					$graphics = [Drawing.Graphics]::FromImage($bitmap)
					$graphics.Clear([Drawing.Color]::Transparent)
					$graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::HighQuality
					$graphics.TextRenderingHint = [Drawing.Text.TextRenderingHint]::AntiAliasGridFit
					$graphics.DrawImage($metafile, [Drawing.Rectangle]::new(0, 0, $pixelsX, $pixelsY))
					$count++
					$pngName = $imagePrefix + $count + '.png'
					$png = $archive.CreateEntry('word/' + $pngName).Open()
					$bitmap.Save($png, [Drawing.Imaging.ImageFormat]::Png)
				} finally {
					if ($png) { $png.Dispose() }; if ($graphics) { $graphics.Dispose() }
					if ($bitmap) { $bitmap.Dispose() }; if ($metafile) { $metafile.Dispose() }
					$memory.Dispose(); $sourceStream.Dispose()
				}
				$rel.SetAttribute('Target', $pngName)
				$picture = $xml.CreateElement('w', 'pict', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
				foreach ($shapeType in $node.SelectNodes('v:shapetype', $ns)) {
					[void]$picture.AppendChild($shapeType.CloneNode($true))
				}
				$staticShape = $shape.CloneNode($true)
				$staticShape.RemoveAttribute('ole', 'urn:schemas-microsoft-com:office:office')
				[void]$picture.AppendChild($staticShape)
				[void]$node.ParentNode.ReplaceChild($picture, $node)
			}
			Write-ZipXml $archive $part $xml
			Write-ZipXml $archive $relsName $rels
		}
		if ($count -gt 0) {
			$types = Read-ZipXml $archive '[Content_Types].xml'
			if (-not @($types.DocumentElement.ChildNodes | Where-Object { $_.LocalName -eq 'Default' -and $_.Extension -eq 'png' }).Count) {
				$type = $types.CreateElement('Default', $types.DocumentElement.NamespaceURI)
				$type.SetAttribute('Extension','png'); $type.SetAttribute('ContentType','image/png')
				[void]$types.DocumentElement.AppendChild($type)
				Write-ZipXml $archive '[Content_Types].xml' $types
			}
		}
	} finally { $archive.Dispose() }
	return $count
}
function Repair-FormulaLayout($Document, $Equation) {
	if (-not $Equation.display) { return }
	$range = Get-BookmarkRange $Document $Equation.bookmark
	$paragraph = $range.Paragraphs.Item(1)
	$format = $paragraph.Format
	# The writer already supplied correct tabs for the applicable text region;
	# remember them before conversion and restore the actual paragraph format.
	if ($script:paragraphFormats.ContainsKey($Equation.id)) {
		$paragraph.Format = $script:paragraphFormats[$Equation.id]
	}
	$format = $paragraph.Format
	if ($Equation.numbered) {
		$format.Alignment = 0
		$format.FirstLineIndent = 0; $format.LeftIndent = 0; $format.RightIndent = 0
		$format.CharacterUnitFirstLineIndent = 0; $format.CharacterUnitLeftIndent = 0; $format.CharacterUnitRightIndent = 0
	}
}
function Get-FormulaState($Document, $Equations) {
	$converted = 0; $retained = 0
	foreach ($equation in $Equations) {
		$range = Get-BookmarkRange $Document $equation.bookmark
		if ($range.InlineShapes.Count -eq 1) {
			$shape = $range.InlineShapes.Item(1)
			if ($shape.Type -ne 1 -or $shape.OLEFormat.ProgID -ne 'Equation.DSMT4') { throw ('Unexpected OLE object at ' + $equation.id) }
			$converted++
		} elseif ($range.InlineShapes.Count -eq 0 -and $range.Text -ceq ('$' + $equation.tex + '$')) { $retained++ }
		else { throw ('Equation integrity check failed at ' + $equation.id) }
	}
	if ($converted -gt 0 -and $retained -gt 0) { throw 'Mixed MathType/TeX output is not permitted for this conversion batch.' }
	return [ordered]@{converted = $converted; retained_tex = $retained}
}
function Get-LayoutEvidence($Document, $Equations, $NumberBookmarks) {
	$items = @()
	foreach ($equation in $Equations) {
		if (-not $equation.display) { continue }
		$range = Get-BookmarkRange $Document $equation.bookmark
		$paragraph = $range.Paragraphs.Item(1)
		$start = $range.Duplicate; $start.Collapse(1)
		$last = $range.Duplicate; $last.Collapse(0)
		$sameLine = $null
		if ($equation.numbered -and $NumberBookmarks.ContainsKey($equation.id)) {
			$number = Get-BookmarkRange $Document $NumberBookmarks[$equation.id]
			$numStart = $number.Duplicate; $numStart.Collapse(1)
			# Word's rendered line/page indices are stronger than a same-paragraph
			# test. PDF inspection is still required for bounds and exact centering.
			$sameLine = ($start.Information(10) -eq $numStart.Information(10) -and
						$start.Information(3) -eq $numStart.Information(3))
		}
		$items += [ordered]@{
			id = $equation.id; page = $start.Information(3)
			start_line = $start.Information(10); end_line = $last.Information(10)
			number_same_line = $sameLine
			same_paragraph = ($range.Paragraphs.Count -eq 1)
			ole_width_pt = $(if ($range.InlineShapes.Count -eq 1) { $range.InlineShapes.Item(1).Width } else { $null })
			centering_and_bounds = 'requires_page_preview'
		}
	}
	return $items
}

$word = $null; $doc = $null; $owned = $false; $script:paragraphFormats = @{}
$report = [ordered]@{
	conversion = 'not_run'; numbering = 'not_run'; compatible = $false
	pdf_exported = $false; saved_and_reopened = $false
	warnings = @(); unverified = @('Per-equation visual centering, bounds, and mathematical boldness',
		'Native numbering format change and final-document move/reference round trip')
	profile = 'MathType 6.9 desktop Word add-in; continuous Arabic (#E1)'
}
$numberBookmarks = @{}
try {
	$context = Get-EnvironmentReport
	if (-not $context.desktop_context) { throw 'This process uses the Codex sandbox account. Run Word automation in the approved desktop-user context; do not copy licensing registry values.' }
	if (-not $context.word_registered) { throw 'Microsoft Word COM is not registered.' }
	$manifest = [IO.File]::ReadAllText($ManifestPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
	$equations = @($manifest.equations)
	$numbered = @($equations | Where-Object { $_.numbered })
	$references = @($manifest.references | Where-Object { $_.kind -eq 'equation' })
	$priorPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
	Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class L2WNativeWindow {
	[DllImport("user32.dll")]
	public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
	$word = New-Object -ComObject Word.Application
	# Word exposes Hwnd on Window, unlike Excel's Application.Hwnd.  A blank
	# document gives this new COM application its own identifiable window.
	$doc = $word.Documents.Add()
	[uint32]$wordPid = 0
	[void][L2WNativeWindow]::GetWindowThreadProcessId([IntPtr]$word.ActiveWindow.Hwnd, [ref]$wordPid)
	if ($wordPid -eq 0 -or $priorPids -contains $wordPid) { throw 'Word did not create an isolated process; refusing to modify an existing instance.' }
	$owned = $true
	$process = Get-Process -Id $wordPid
	Write-JsonFile $ProcessPath @{pid = $wordPid; started_ticks = $process.StartTime.ToUniversalTime().Ticks.ToString()}
	$word.Visible = $false; $word.DisplayAlerts = 0
	$report.word_version = [string]$word.Version
	$report.identity = $context.identity
	Close-Document $doc; $doc = $null
	$doc = Open-Document $word $InputPath
	$initialized = $false
	try {
		$report.template = Initialize-MathType $word
		$initialized = $true
	} catch {
		$report.warnings += $_.Exception.Message
		$report.conversion = 'tex_retained'; $report.numbering = 'unavailable'
	}
	# Validate the entire input before any mutation; preserve exact body ranges.
	foreach ($equation in $equations) {
		$range = Get-BookmarkRange $doc $equation.bookmark
		if ($range.Text -cne ('$' + $equation.tex + '$')) { throw ('Equation bookmark does not cover its exact $TeX$ body: ' + $equation.id) }
		if ($equation.display) { $script:paragraphFormats[$equation.id] = $range.Paragraphs.Item(1).Format.Duplicate }
	}
	if ($doc.OMaths.Count -gt 0) { throw 'Input contains unexpected OMML; use literal TeX bookmarks.' }
	if ($numbered.Count -eq 0 -and $references.Count -eq 0) { $report.numbering = 'not_needed' }
	elseif ($initialized) {
		try {
			$report.numbering_probe = Test-NativeNumbering $word
			$doc.Activate()
			Set-NumberFormat $word $doc
			$index = 0; $labelToNumber = @{}
			foreach ($equation in $numbered) {
				$index++
				$name = 'L2WNum' + $index.ToString('0000')
				$range = Get-BookmarkRange $doc $equation.bookmark
				$marker = Find-Text $range.Paragraphs.Item(1).Range ('L2W_NUMBER_' + $equation.id)
				[void](Insert-NativeNumber $word $doc $marker $name)
				$numberBookmarks[$equation.id] = $name
				if ($equation.label) { $labelToNumber[$equation.label] = $name }
			}
			foreach ($reference in $references) {
				$placeholder = Get-BookmarkRange $doc $reference.bookmark
				$target = Get-BookmarkRange $doc $labelToNumber[$reference.target]
				Insert-NativeReference $word $doc $placeholder $target
			}
			Update-NativeNumbers $word $doc
			$report.numbering = 'native_created'
		} catch {
			$report.warnings += ('Native numbering rolled back: ' + $_.Exception.Message + ' [' + $_.ScriptStackTrace + ']')
			Close-Document $doc; $doc = Open-Document $word $InputPath
			$numberBookmarks = @{}
			$report.numbering = 'unavailable'
		}
	}
	# This checkpoint separates native numbering from all-or-nothing TeX conversion.
	$checkpoint = Join-Path ([IO.Path]::GetDirectoryName($OutputPath)) 'numbered-tex.docx'
	Save-DocumentAs $doc $checkpoint
	if ($equations.Count -eq 0) { $report.conversion = 'not_needed' }
	elseif ($initialized) {
		try {
			$button = 'MathType_B_TeXToggle'
			$disabled = $word.Run('GetDisabledString', [ref]$button)
			if ($disabled) { throw ('TeX toggle disabled: ' + $disabled) }
			foreach ($equation in $equations) {
				$range = Get-BookmarkRange $doc $equation.bookmark
				$position = $range.Start
				$story = $range.Duplicate; $story.WholeStory()
				$previousCount = $story.InlineShapes.Count
				$range.Select()
				[void]$word.Run('MTCommand_TeXToggle')
				if ($story.InlineShapes.Count -ne ($previousCount + 1)) { throw ('TeX conversion did not create exactly one object at ' + $equation.id) }
				$objectRange = $null
				for ($i = 1; $i -le $story.InlineShapes.Count; $i++) {
					$shape = $story.InlineShapes.Item($i)
					if ($shape.Range.Start -eq $position -and $shape.OLEFormat.ProgID -eq 'Equation.DSMT4') {
						# An OLE EMBED field occupies more than one Word character
						# position. Preserve its complete Range, not [start,start+1].
						$objectRange = $shape.Range.Duplicate
						break
					}
				}
				if (-not $objectRange) { throw ('TeX conversion did not produce Equation.DSMT4 at ' + $equation.id) }
				[void]$doc.Bookmarks.Add($equation.bookmark, $objectRange)
				Repair-FormulaLayout $doc $equation
			}
			$report.formulas = Get-FormulaState $doc $equations
			$report.conversion = 'converted'
		} catch {
			$report.warnings += ('All formula bodies restored to TeX: ' + $_.Exception.Message + ' [' + $_.ScriptStackTrace + ']')
			Close-Document $doc; $doc = Open-Document $word $checkpoint
			$report.conversion = 'tex_retained'
		}
	}
	if ($report.numbering -eq 'native_created') { Update-NativeNumbers $word $doc }
	Update-AllStories $doc
	$doc.Repaginate()
	$report.formulas = Get-FormulaState $doc $equations
	$report.view_settings_applied = Set-TargetView $doc
	Save-DocumentAs $doc $OutputPath
	Close-Document $doc; $doc = Open-Document $word $OutputPath
	if ($report.numbering -eq 'native_created') { Update-NativeNumbers $word $doc }
	$report.formulas = Get-FormulaState $doc $equations
	$report.saved_and_reopened = $true
	$report.view_settings_applied = Set-TargetView $doc
	$finalView = $doc.ActiveWindow.View
	$report.view_settings = [ordered]@{
		show_all = [bool]$finalView.ShowAll
		show_paragraphs = [bool]$finalView.ShowParagraphs
		show_field_codes = [bool]$finalView.ShowFieldCodes
	}
	if (-not $report.view_settings_applied) {
		$readback = $report.view_settings | ConvertTo-Json -Compress
		$report.warnings += ('Document view switches were requested, but Word did not confirm all disabled: ' + $readback + '. Check nonprinting marks in the Word editing view manually.')
		$report.unverified += 'Editing-view nonprinting marks: this Word instance did not confirm the requested view switches; inspect view_settings and verify manually.'
	}
	$doc.Repaginate()
	$report.layout = @(Get-LayoutEvidence $doc $equations $numberBookmarks)
	foreach ($layout in $report.layout) {
		if ($layout.number_same_line -eq $false -or $layout.start_line -ne $layout.end_line) {
			$report.warnings += ('Rendered formula/number wraps onto another line: ' + $layout.id)
		}
	}
	$report.pages = $doc.ComputeStatistics(2)
	if ($ExportPdf) {
		$renderDocument = $null
		try {
			$pdf = [IO.Path]::ChangeExtension($OutputPath, '.pdf')
			$doc.Save()
			$renderPath = Join-Path ([IO.Path]::GetDirectoryName($OutputPath)) 'pdf-rendering-copy.docx'
			$report.pdf_rasterized_equations = New-PdfRenderingCopy $OutputPath $renderPath
			$renderDocument = Open-Document $word $renderPath
			[void](Set-TargetView $renderDocument)
			$renderDocument.Repaginate()
			$renderDocument.ExportAsFixedFormat($pdf, 17)
			$report.pdf_exported = Test-Path -LiteralPath $pdf
			$report.pdf_method = '600dpi rasterized MathType previews in a disposable copy; primary DOCX retains native OLE'
			$report.pdf_equations_searchable = $false
		} catch { $report.warnings += ('PDF export failed: ' + $_.Exception.Message) }
		finally { Close-Document $renderDocument; $doc.Activate() }
	}
	$doc.Save()
} catch {
	$report.fatal_error = $_.Exception.Message
} finally {
	try { Close-Document $doc } catch { $report.warnings += ('Document cleanup: ' + $_.Exception.Message) }
	if ($word) {
		if ($owned) {
			try { $save = 0; $word.Quit([ref]$save) } catch { $report.warnings += ('Word cleanup: ' + $_.Exception.Message) }
		}
		Release-Com $word
	}
	[GC]::Collect(); [GC]::WaitForPendingFinalizers()
	Write-JsonFile $ResultPath $report
}
if ($report.Contains('fatal_error')) { exit 1 }
