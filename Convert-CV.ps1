<#
.SYNOPSIS
    Generates a PDF of the CV using a headless browser.

.DESCRIPTION
    Uses Microsoft Edge (or Chrome) in headless mode to print index.html to a
    pixel-perfect PDF. Requires Edge or Chrome to be installed.

.PARAMETER HtmlPath
    Path to the source HTML file. Defaults to index.html in the script folder.

.PARAMETER OutputBaseName
    Base filename for the output PDF.

.PARAMETER OutputDirectory
    Directory to write the PDF. Defaults to the script folder.

.EXAMPLE
    .\Convert-CV.ps1
#>
[CmdletBinding()]
param(
    [string]$HtmlPath       = 'index.html',
    [string]$OutputBaseName = 'Nathan-Gregory-CV',
    [string]$OutputDirectory = '.'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

function Resolve-CvPath {
    param([string]$Path, [string]$Base)
    if ([System.IO.Path]::IsPathRooted($Path)) { return [System.IO.Path]::GetFullPath($Path) }
    return [System.IO.Path]::GetFullPath((Join-Path $Base $Path))
}

$resolvedHtml   = Resolve-CvPath -Path $HtmlPath       -Base $scriptDir
$resolvedOutDir = Resolve-CvPath -Path $OutputDirectory -Base $scriptDir

if (-not (Test-Path $resolvedHtml))   { throw "HTML not found: $resolvedHtml" }
if (-not (Test-Path $resolvedOutDir)) { New-Item -ItemType Directory -Path $resolvedOutDir | Out-Null }

$pdfPath = Join-Path $resolvedOutDir ($OutputBaseName + '.pdf')
$fileUri = 'file:///' + $resolvedHtml.Replace('\', '/')

$browsers = @(
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
)
$browser = $browsers | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $browser) { throw 'No supported browser found. Install Microsoft Edge or Google Chrome.' }

Write-Host "Browser : $(Split-Path $browser -Leaf)"
Write-Host "Source  : $resolvedHtml"
Write-Host "PDF     : $pdfPath"

$tempProfile = Join-Path $env:TEMP ('cv-export-' + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $tempProfile | Out-Null

try {
    $args = @(
        '--headless=old',
        '--disable-gpu',
        '--no-pdf-header-footer',
        '--run-all-compositor-stages-before-draw',
        "--print-to-pdf=$pdfPath",
        "--user-data-dir=$tempProfile"
        $fileUri
    )
    $proc = Start-Process -FilePath $browser -ArgumentList $args -Wait -PassThru -NoNewWindow
    if (Test-Path $pdfPath) {
        Write-Host 'Saved PDF.' -ForegroundColor Green
    } else {
        Write-Warning 'PDF was not created. Exit code: ' + $proc.ExitCode
    }
}
finally {
    Remove-Item $tempProfile -Recurse -Force -ErrorAction SilentlyContinue
}
