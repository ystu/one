<#
.SYNOPSIS
Format an existing three-level speech outline DOCX without changing its content.
.DESCRIPTION
Requires a simple DOCX with title, author, then native Word numbering at levels
0, 1 and 2. Preserves fonts, sizes, numbering, all text and other package parts.
Sets exact line heights and disables paragraph grid snapping. The stored grid
remains intact. No Word, Python or external package is needed for formatting.
Estimated height is a warning only; verify actual pagination and a page preview.
Output is written atomically; a locked destination fails without force-closing Word.
.EXAMPLE
.\scripts\format_speech_outline.ps1 -InputPath source.docx -OutputPath output.docx -NarrowMargins
.EXAMPLE
.\scripts\format_speech_outline.ps1 -InputPath outline.docx -OutputPath outline.docx -LineSpacingPt 21.5 -NarrowMargins
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [ValidateRange(18, 30)][double]$LineSpacingPt = 21.5,
    [switch]$NarrowMargins
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new()
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$uri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$src = (Resolve-Path -LiteralPath $InputPath).Path
$dst = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if ([IO.Path]::GetExtension($src) -ne '.docx' -or [IO.Path]::GetExtension($dst) -ne '.docx') {
    throw 'InputPath and OutputPath must be DOCX files.'
}
$parent = [IO.Path]::GetDirectoryName($dst)
if (-not [IO.Directory]::Exists($parent)) { throw 'Output directory does not exist.' }
$staging = Join-Path $parent ('.outline-' + [Guid]::NewGuid().ToString('N') + '.docx')

function Set-WordAttribute($Node, [string]$Name, $Value) {
    [void]$Node.SetAttribute($Name, $uri, [string]$Value)
}
function Ensure-Child($Parent, [string]$Name) {
    $node = $Parent.SelectSingleNode("w:$Name", $ns)
    if (-not $node) {
        $node = $xml.CreateElement('w', $Name, $uri)
        [void]$Parent.AppendChild($node)
    }
    return ,$node
}

try {
    [IO.File]::Copy($src, $staging)
    $zip = [IO.Compression.ZipFile]::Open($staging, 'Update')
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if (-not $entry) { throw 'Missing word/document.xml.' }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('w', $uri)
        $body = $xml.SelectSingleNode('/w:document/w:body', $ns)
        if ($body.SelectNodes('w:tbl | .//w:sectPr', $ns).Count -ne 1 -or
            $body.SelectNodes('w:tbl', $ns).Count -gt 0 -or
            $body.SelectNodes('w:sdt | w:altChunk', $ns).Count -gt 0) {
            throw 'Only simple outlines with one section and no tables/content controls are supported.'
        }
        $paragraphs = @($body.SelectNodes('w:p', $ns))
        if ($paragraphs.Count -lt 5) { throw 'Expected title, author and three-level outline paragraphs.' }
        $levels = @()
        foreach ($para in $paragraphs | Select-Object -Skip 2) {
            $level = $para.SelectSingleNode('w:pPr/w:numPr/w:ilvl', $ns)
            $id = $para.SelectSingleNode('w:pPr/w:numPr/w:numId', $ns)
            if (-not $level -or -not $id -or $level.GetAttribute('val', $uri) -notmatch '^[012]$') {
                throw 'Every body paragraph must use native three-level Word numbering. No content was removed.'
            }
            $levels += [int]$level.GetAttribute('val', $uri)
        }
        if (@($levels | Sort-Object -Unique).Count -ne 3) { throw 'All three numbering levels must be present.' }
        $beforeText = @($xml.SelectNodes('//w:t', $ns) | ForEach-Object InnerText) -join [char]0
        $beforeNumbers = @($xml.SelectNodes('//w:numPr', $ns) | ForEach-Object OuterXml) -join [char]0
        $majorCount = 0
        for ($i = 0; $i -lt $paragraphs.Count; $i++) {
            $pp = $paragraphs[$i].SelectSingleNode('w:pPr', $ns)
            if (-not $pp) {
                $pp = $xml.CreateElement('w', 'pPr', $uri)
                [void]$paragraphs[$i].PrependChild($pp)
            }
            $spacing = Ensure-Child $pp 'spacing'
            foreach ($a in @('before', 'after')) { Set-WordAttribute $spacing $a 0 }
            foreach ($a in @('beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing')) {
                $spacing.RemoveAttribute($a, $uri)
            }
            $height = if ($i -eq 0) { 28 } elseif ($i -eq 1) { 18 } else { $LineSpacingPt }
            Set-WordAttribute $spacing 'line' ([int][Math]::Round($height * 20))
            Set-WordAttribute $spacing 'lineRule' 'exact'
            $snap = Ensure-Child $pp 'snapToGrid'
            Set-WordAttribute $snap 'val' 0
            [void]$pp.InsertBefore($snap, $spacing)
            if ($i -ge 2 -and $levels[$i - 2] -eq 0) {
                $majorCount++
                if ($majorCount -gt 1) { Set-WordAttribute $spacing 'before' 60 }
            }
        }
        $section = $body.SelectSingleNode('w:sectPr', $ns)
        $size = $section.SelectSingleNode('w:pgSz', $ns)
        if ($size.GetAttribute('w', $uri) -ne '11906' -or $size.GetAttribute('h', $uri) -ne '16838') {
            throw 'Expected portrait A4. Review page settings before applying this preset.'
        }
        $margins = $section.SelectSingleNode('w:pgMar', $ns)
        if (-not $margins) { throw 'Missing explicit page margins.' }
        if ($NarrowMargins) {
            foreach ($a in @('top', 'bottom', 'left', 'right')) { Set-WordAttribute $margins $a 720 }
        }
        $afterText = @($xml.SelectNodes('//w:t', $ns) | ForEach-Object InnerText) -join [char]0
        $afterNumbers = @($xml.SelectNodes('//w:numPr', $ns) | ForEach-Object OuterXml) -join [char]0
        if ($beforeText -cne $afterText -or $beforeNumbers -cne $afterNumbers) { throw 'Content or numbering changed.' }
        $availablePt = (16838 - [int]$margins.GetAttribute('top', $uri) - [int]$margins.GetAttribute('bottom', $uri)) / 20
        $minimumPt = 46 + ($paragraphs.Count - 2) * $LineSpacingPt + [Math]::Max(0, $majorCount - 1) * 3
        $entry.Delete()
        $entry = $zip.CreateEntry('word/document.xml')
        $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
        try { $writer.Write($xml.OuterXml) } finally { $writer.Dispose() }
    } finally { $zip.Dispose() }

    # Check parseability of every XML part before replacing the destination.
    $check = [IO.Compression.ZipFile]::OpenRead($staging)
    try {
        foreach ($part in $check.Entries | Where-Object { $_.FullName -match '\.(xml|rels)$' }) {
            $reader = [IO.StreamReader]::new($part.Open())
            try { [xml]$parsed = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
    } finally { $check.Dispose() }
    if ([IO.File]::Exists($dst)) { [IO.File]::Replace($staging, $dst, [System.Management.Automation.Language.NullString]::Value) }
    else { [IO.File]::Move($staging, $dst) }
    [pscustomobject]@{
        OutputPath = $dst
        Paragraphs = $paragraphs.Count
        BodyLineSpacingPt = $LineSpacingPt
        MinimumContentHeightPt = $minimumPt
        AvailableHeightPt = $availablePt
        TextAndNumberingPreserved = $true
        PaginationVerified = $false
        NextStep = 'Verify actual page count and preview; wrapping can increase height.'
    }
    if ($minimumPt -gt $availablePt) {
        Write-Warning 'Even without wrapped lines this exceeds one page. Adjust layout or discuss content with the user.'
    }
} finally {
    if ([IO.File]::Exists($staging)) { [IO.File]::Delete($staging) }
}
