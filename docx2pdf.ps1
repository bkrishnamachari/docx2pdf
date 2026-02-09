<#
.SYNOPSIS
    Convert a DOCX file to PDF using Microsoft Word COM automation.

.DESCRIPTION
    This script opens a .docx file in Microsoft Word (via COM), exports it as
    a PDF, and optionally opens the resulting PDF. Requires Microsoft Word to
    be installed on the system.

.PARAMETER InputPath
    Path to the input .docx file.

.PARAMETER OutputPath
    (Optional) Path for the output .pdf file. Defaults to the same directory
    and name as the input file with a .pdf extension.

.PARAMETER Open
    (Optional) If specified, opens the resulting PDF after conversion.

.EXAMPLE
    .\docx2pdf.ps1 -InputPath "C:\docs\report.docx"

.EXAMPLE
    .\docx2pdf.ps1 -InputPath "report.docx" -OutputPath "output.pdf" -Open
#>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputPath,

    [Parameter(Mandatory=$false, Position=1)]
    [string]$OutputPath,

    [switch]$Open
)

$ErrorActionPreference = "Stop"

# Resolve full paths
$InputPath = (Resolve-Path $InputPath).Path

if (-not $OutputPath) {
    $OutputPath = [System.IO.Path]::ChangeExtension($InputPath, ".pdf")
} else {
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
}

if (-not (Test-Path $InputPath)) {
    Write-Error "Input file not found: $InputPath"
    exit 1
}

Write-Host "Converting: $InputPath"
Write-Host "       To: $OutputPath"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($InputPath)
    # 17 = wdExportFormatPDF
    $doc.ExportAsFixedFormat($OutputPath, 17, $false, 0, 0, 0, 0, 0, $false, $false, 0, $false, $true, $false)
    $doc.Close($false)
    Write-Host "Conversion successful."
} catch {
    Write-Error "Conversion failed: $($_.Exception.Message)"
    exit 1
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

if ($Open) {
    Start-Process $OutputPath
}
