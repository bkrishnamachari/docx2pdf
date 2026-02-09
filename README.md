# docx2pdf

A simple CLI tool to convert DOCX files to PDF using Microsoft Word COM automation on Windows.

## Requirements

- Windows with Microsoft Word installed
- PowerShell 5.1+

## Usage

### Standalone PowerShell Script

```powershell
# Basic conversion (PDF saved alongside the DOCX)
.\docx2pdf.ps1 -InputPath "document.docx"

# Specify output path
.\docx2pdf.ps1 -InputPath "document.docx" -OutputPath "output.pdf"

# Convert and open the PDF
.\docx2pdf.ps1 -InputPath "document.docx" -Open
```

### Claude Code Skill

The `docx2pdf.md` file can be used as a [Claude Code](https://claude.ai/claude-code) custom slash command. Copy it to `~/.claude/commands/` to enable the `/docx2pdf` command:

```
cp docx2pdf.md ~/.claude/commands/
```

Then in Claude Code:

```
/docx2pdf path/to/document.docx
```

## How It Works

The tool uses Microsoft Word's COM automation interface to open a DOCX file and export it as PDF. This produces high-fidelity output since Word itself handles the rendering, preserving all formatting, fonts, and layout exactly as Word would print them.

## License

PolyForm Noncommercial License 1.0.0 - See [LICENSE.md](LICENSE.md)
