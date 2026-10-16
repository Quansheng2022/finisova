# Finisova 2.1.0 — Build5 Internal Evaluation

**Status: Internal Evaluation Only — Not for Public Distribution**

Finisova transforms Markdown into professionally formatted Microsoft Word documents and supports PDF export through Microsoft Word.

## Key Features

- Markdown to editable DOCX documents
- PDF export through Microsoft Word
- DOCX, PDF, and Both output modes
- Graphical and command-line interfaces
- Custom style presets, profiles, and Word templates
- Mathematical equations, tables, captions, cross-references, and bibliographies
- Mermaid diagrams and advanced document preview
- Batch conversion and Project Mode
- Local document processing

## Installation Requirements

- Windows desktop environment
- Microsoft Word is not required for DOCX generation
- Microsoft Word is required for PDF export
- LibreOffice PDF export is not supported in v2.1.0

## Installer

**Filename:** `Finisova_v2.1.0_Setup.exe`

**Size:** 114,101,262 bytes

**SHA-256:**

`2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d`

Verify the installer in PowerShell:

```powershell
Get-FileHash .\Finisova_v2.1.0_Setup.exe -Algorithm SHA256
```

## Quick Start — CLI

After installation:

```powershell
$cli = "$env:LOCALAPPDATA\Programs\Finisova\cli\finisova-cli.exe"

# Display help
& $cli --help

# Convert Markdown to DOCX
& $cli ".\example.md" --format docx --no-open

# Convert Markdown to PDF (requires Microsoft Word)
& $cli ".\example.md" --format pdf --no-open

# Generate DOCX and PDF
& $cli ".\example.md" --format both --no-open
```

## Quick Start — GUI

1. Launch Finisova.
2. Select a Markdown file.
3. Choose DOCX, PDF, or Both.
4. Select an output folder and optional style preset.
5. Preview the document if needed.
6. Click Convert and verify the output.

## Documentation

- [Product README](../../README.md)
- [English User Guide](../../docs/Finisova_2.1.0_User_Guide_en.md)
- [Chinese User Guide](../../docs/Finisova_2.1.0_User_Guide_zh-CN.md)

Documentation links require the corresponding files to exist in the repository.

## Known Limitations

- Clean Windows installation and uninstall verification remain incomplete.
- Microsoft Word is required for PDF generation.
- CLI style-preset selection is not yet implemented.
- Preview is not a pixel-perfect Word rendering.
- End-to-end SVG embedding in DOCX remains unverified for Build5.
- Installed offline preview and Word-absent behavior remain to be qualified.
- Third-party licensing and redistribution compliance review remains open.

## Release Status

Build5 passed its recorded internal packaging checks and developer-machine GUI/CLI smoke tests.

This build is **not** a release candidate or a production release. It remains subject to separately authorized Windows installation QA, licensing review, and final release approval.

**Publisher label:** Arculoric  
**Product:** Finisova  
**Version:** 2.1.0