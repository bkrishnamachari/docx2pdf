# docx2pdf

A simple Windows CLI tool that converts DOCX files to PDF using Microsoft Word's built-in PDF export, then opens the result.

## Requirements

- Windows
- Microsoft Word (uses COM automation)
- PowerShell

## Installation

1. Clone or download this repo
2. Copy `docx2pdf.bat` and `docx2pdf.ps1` to a directory in your PATH (e.g., `%USERPROFILE%\bin`)

```cmd
git clone https://github.com/bkrishnamachari/docx2pdf.git
copy docx2pdf\docx2pdf.bat %USERPROFILE%\bin\
copy docx2pdf\docx2pdf.ps1 %USERPROFILE%\bin\
```

## Usage

```cmd
docx2pdf report.docx                # converts to report.pdf and opens it
docx2pdf report.docx output.pdf     # converts to specific output path and opens it
```

## How it works

The tool uses Microsoft Word's COM automation interface via PowerShell to open the DOCX file and save it as PDF (format code 17). This produces identical output to using Word's "Save as PDF" manually. After conversion, the PDF is automatically opened with your default PDF viewer.

The COM objects are properly released after conversion to prevent memory leaks.

## Claude Code Skill

A Claude Code slash command is included in `claude-skill/docx2pdf.md`. To install it:

```cmd
copy claude-skill\docx2pdf.md %USERPROFILE%\.claude\commands\
```

Then use `/docx2pdf path\to\file.docx` from any Claude Code session.

## License

PolyForm Noncommercial 1.0.0 - see [LICENSE](LICENSE) for details.
