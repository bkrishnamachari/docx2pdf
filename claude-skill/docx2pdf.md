---
argument-hint: <path-to-docx-file>
description: Convert a DOCX file to PDF using Word COM automation and open it
allowed-tools: Bash, Read, Write
---

Convert the specified DOCX file to PDF using Microsoft Word COM automation via PowerShell, then open the resulting PDF.

## Instructions

1. The user wants to convert a DOCX file to PDF: `$ARGUMENTS`
2. Resolve the full absolute path of the input DOCX file.
3. If no output path is specified, save the PDF in the same directory with the same name but `.pdf` extension.
4. Use this PowerShell approach to convert (write a temp .ps1 script and execute it):

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open("<absolute-docx-path>")
    $doc.SaveAs("<absolute-pdf-path>", 17)
    $doc.Close($false)
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
```

5. Execute the script via: `powershell -ExecutionPolicy Bypass -File "<temp-script-path>"`
6. Clean up the temp script after execution.
7. Open the resulting PDF with: `start "" "<pdf-path>"`
8. Report success or failure to the user.

IMPORTANT: Use double backslashes in the PowerShell paths. Set a 60-second timeout on the conversion command. The `start` command MUST use the `start "" "path"` format (empty string for window title).
