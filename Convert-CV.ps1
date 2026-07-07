<#
.SYNOPSIS
    Converts index.html to .docx and .pdf using Microsoft Word.

.DESCRIPTION
    Requires Microsoft Word to be installed.
    Reads index.html, saves as DOCX then exports to PDF.

.PARAMETER HtmlPath
    Path to the source HTML file. Defaults to index.html in the script directory.

.PARAMETER OutputBaseName
    Base filename (no extension) for output files. Defaults to Nathan-Gregory-CV.

.PARAMETER OutputDirectory
    Directory to save output files. Defaults to the script directory.

.EXAMPLE
    .\Convert-CV.ps1

.EXAMPLE
    .\Convert-CV.ps1 -HtmlPath index.html -OutputBaseName "Nathan-Gregory-CV" -OutputDirectory .
#>

[CmdletBinding()]
param(
    [string]$HtmlPath       = $null,
    [string]$OutputBaseName = 'Nathan-Gregory-CV',
    [string]$OutputDirectory = $null
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot

if (-not $HtmlPath) {
    $HtmlPath = Join-Path $scriptDir 'index.html'
}

if (-not $OutputDirectory) {
    $OutputDirectory = $scriptDir
}

$HtmlPath       = [System.IO.Path]::GetFullPath($HtmlPath)
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)

if (-not (Test-Path $HtmlPath)) {
    Write-Error "Source HTML not found: $HtmlPath"
    exit 1
}

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
}

$docxPath = Join-Path $OutputDirectory "$OutputBaseName.docx"
$pdfPath  = Join-Path $OutputDirectory "$OutputBaseName.pdf"

Write-Host "Source : $HtmlPath"
Write-Host "DOCX   : $docxPath"
Write-Host "PDF    : $pdfPath"

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    # Open HTML as Word document
    $doc = $word.Documents.Open(
        $HtmlPath,
        $false,  # ConfirmConversions
        $false,  # ReadOnly
        $false   # AddToRecentFiles
    )

    if ($null -eq $doc) {
        throw "Word failed to open: $HtmlPath"
    }

    # Set page size to A4 and apply narrow margins (0.5 inch = 720 twips)
    try {
        foreach ($section in $doc.Sections) {
            $pm = $section.PageSetup
            $pm.PaperSize    = 9      # wdPaperA4
            $pm.Orientation  = 0      # wdOrientPortrait
            $pm.TopMargin    = 720
            $pm.BottomMargin = 720
            $pm.LeftMargin   = 720
            $pm.RightMargin  = 720
        }
    }
    catch {
        Write-Warning 'Could not apply custom margins - continuing without them.'
    }

    # Save as DOCX (16 = wdFormatDocumentDefault), suppress compatibility dialog
    $word.DisplayAlerts = 0   # wdAlertsNone
    $doc.SaveAs2([ref]$docxPath, [ref]16)
    Write-Host 'Saved DOCX.'

    # Export as PDF (17 = wdExportFormatPDF)
    $doc.ExportAsFixedFormat($pdfPath, 17)
    Write-Host 'Saved PDF.'

    $doc.Close($false)
    Write-Host 'Done.' -ForegroundColor Green
}
catch {
    Write-Error ('Conversion failed: ' + $_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
