param(
    [Parameter(Position=0)]
    [string]$DocxPath,

    [Parameter(Position=1)]
    [string]$PdfPath
)

if (-not $DocxPath) {
    Write-Host "Usage: docx2pdf <file.docx> [output.pdf]"
    Write-Host ""
    Write-Host "Converts a DOCX file to PDF using Microsoft Word, then opens it."
    Write-Host "If no output path is given, saves PDF next to the DOCX file."
    exit 1
}

$DocxPath = (Resolve-Path $DocxPath -ErrorAction SilentlyContinue).Path
if (-not $DocxPath -or -not (Test-Path $DocxPath)) {
    Write-Host "Error: File not found: $DocxPath"
    exit 1
}

if ([IO.Path]::GetExtension($DocxPath) -ne ".docx") {
    Write-Host "Error: Input file must be a .docx file"
    exit 1
}

if (-not $PdfPath) {
    $PdfPath = [IO.Path]::ChangeExtension($DocxPath, ".pdf")
} else {
    $PdfPath = [IO.Path]::GetFullPath($PdfPath)
}

Write-Host "Converting: $DocxPath"
Write-Host "       To: $PdfPath"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($DocxPath)
    $doc.SaveAs($PdfPath, 17)
    $doc.Close($false)
    Write-Host "Conversion complete."
} catch {
    Write-Host "Error: $_"
    Write-Host "Is Microsoft Word installed?"
    exit 1
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Opening: $PdfPath"
Start-Process $PdfPath
