# Finisova 2.1.0

**Turn Markdown into polished Word documents — locally, privately, and without a subscription.**

Finisova is a Windows desktop application for turning Markdown content into professionally formatted Microsoft Word documents (`.docx`) and, when Microsoft Word is available, PDFs. It combines a graphical interface, a command-line converter, reusable styles, document-oriented features, and a local preview.

> **Release status — important:** Version **2.1.0 Build5** is a **frozen internal evaluation candidate**, not a publicly released product. Its executable and installer passed the recorded build and developer-machine smoke checks; **clean-machine installation, Word-absent behavior, installed offline preview, and commercial third-party license clearance remain open**. Do not treat this README as approval to publish or distribute the installer.

**Product:** Finisova · **Version:** 2.1.0 · **Platform:** Windows desktop · **Packaged publisher label:** Arculoric

[Features](#features) · [Getting started](#getting-started) · [CLI](#command-line-interface) · [Requirements](#system-requirements) · [Limitations](#known-limitations-and-verification-status) · [Troubleshooting](#troubleshooting) · [Licensing](#licensing-and-distribution)

## Why Finisova?

Markdown is convenient for writing technical notes, reports, documentation, AI-generated drafts, and research material. Microsoft Word is often more convenient for sharing, editing, reviewing, and delivering finalized documents.

Finisova helps bridge that workflow:

- **Write in Markdown** using your editor of choice.
- **Produce editable Word documents** with structured headings, tables, equations, and document formatting.
- **Choose a presentation** using profiles, templates, and supported reusable style presets.
- **Export to PDF** through Microsoft Word's native PDF export when Word is installed and available.
- **Work on your own computer** rather than uploading documents to a hosted conversion service.

The goal is to minimize time spent manually correcting the output in Word. Exact page layout and visual fidelity still depend on the document content, fonts, Word environment, and supported feature subset.

## Features

The following describes supported product capability as recorded by the v2.0/v2.1 engineering program. It is **not** a claim that every capability has passed external installed-Windows acceptance testing.

| Area | Capability | Notes |
|---|---|---|
| Core conversion | Markdown → DOCX | DOCX generation does **not** require Microsoft Word. |
| PDF output | Markdown → PDF; DOCX **and** PDF (`Both`) | Uses a DOCX-first pipeline and Microsoft Word's native PDF exporter. |
| GUI | File selection, conversion status, output-format selection | Supports `DOCX`, `PDF`, and `Both`; reports errors and partial results where applicable. |
| CLI | Single-file and directory-based workflows | `--format {docx,pdf,both}`, `--output`, `--no-open`, and other options; see `--help`. |
| Formatting | Profiles, Word templates, reusable style presets | GUI preset selection, editor, import/export and reset/inherit workflows. |
| Document structure | Headings, tables, TOC and metadata | Word-native document structures where supported. |
| Academic/technical writing | Editable equations, captions, cross-references, citations and bibliography | Supported syntax and rendering rules are bounded, not universal Markdown extensions. |
| Longer workflows | Batch conversion and Project Mode | Availability and behavior depend on the selected workflow and project configuration. |
| Preview | Formulas, Mermaid, links, rich content and style-aware presentation | Intended to run with local resources; not a pixel-perfect Word/PDF preview. |
| Diagnostics | Validation, errors, and output feedback | A failed or partial conversion should not be interpreted as a complete success. |

### Editable Word output

Finisova's DOCX renderer is designed to retain supported Word-native document structures. In particular, supported mathematical expressions are represented as editable Word equations rather than merely flattened screenshots.

The exact behavior of custom templates, formulas, table layout, diagrams, citations, bookmarks, and cross-references can vary with the source syntax and selected presentation settings. Validate important documents in Word before delivery.

### Styling and templates

Finisova 2.1 adds user-defined presentation presets on top of existing profile/template support. In the GUI, supported workflows include selecting, editing, saving, importing, exporting, and resetting presets.

Presets apply only to supported formatting properties; unspecified properties inherit the existing presentation configuration. Presets should not be understood as unrestricted Word theme editing, arbitrary CSS support, or a full page-layout editor.

**CLI limitation:** CLI `--style-config` / preset selection was **not completed** for the accepted 2.1.0 CLI scope. Use the GUI for preset management, and check `finisova-cli.exe --help` for the actual CLI options available.

### Preview versus final output

Finisova provides a **semantic preview**, not Word's pagination or rendering engine. The preview can help inspect supported headings, formatted text, formulas, Mermaid diagrams, links, and styles, but differences from the final DOCX/PDF may include:

- Page breaks, line wrapping, and font metrics.
- Word-generated fields, page numbers, and table-of-contents updates.
- Complex table layout and image/diagram rendering.
- Unsupported formula, Mermaid, or custom-style constructs.

For final delivery, use the generated document as the authoritative artifact.

## System requirements

| Requirement | DOCX conversion | PDF / Both output |
|---|---|---|
| Operating system | Windows desktop; Windows 11 is the intended clean-machine QA target | Windows desktop with a supported Microsoft Word installation |
| Microsoft Word | **Not required** to generate DOCX | **Required** for native PDF export |
| Python | Not expected to be needed when using the packaged installer | Not expected to be needed when using the packaged installer |
| Internet connection | Local workflow is designed not to require a cloud document-upload service | PDF export is local to Microsoft Word; installed offline behavior is still subject to QA |
| Fonts | Available fonts affect rendering | Word-exported PDF fidelity also depends on fonts and the Word environment |

**Compatibility qualification:** The Build5 executable has passed a developer-host GUI launch/close and CLI DOCX conversion smoke test, but **installation and complete feature behavior on a fresh Windows computer have not yet been qualified**. Specific Windows versions, Office editions, CPU architectures, and hardware requirements must be confirmed before a public support matrix is published.

**LibreOffice is not a PDF backend in 2.1.0.** This option was explicitly deferred; installing LibreOffice alone will not satisfy Finisova's Microsoft Word PDF requirement.

## Getting started

> These instructions are for **authorized internal testers** using an approved disposable or otherwise explicitly designated test computer. There is no approved public Build5 download or release at present.

### 1. Verify the installer

The frozen Build5 internal installer is:

`Finisova_v2.1.0_Setup.exe`

Its recorded SHA-256 is:

```text
2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d
```

In PowerShell:

```powershell
Get-FileHash ".\Finisova_v2.1.0_Setup.exe" -Algorithm SHA256
```

Ensure the hash matches exactly. If you received multiple `.7z.001`, `.7z.002`, etc. volumes, extract them together using 7-Zip before checking the resulting installer.

### 2. Install on an authorized test computer

Run the verified installer. The currently recorded packaging configuration uses a **per-user** installation directory:

```text
%LOCALAPPDATA%\Programs\Finisova
```

The expected application and command-line programs are:

```text
%LOCALAPPDATA%\Programs\Finisova\Finisova.exe
%LOCALAPPDATA%\Programs\Finisova\cli\finisova-cli.exe
```

The installer is designed to provide Start Menu shortcuts and an uninstaller. It does **not** add the CLI directory to your system `PATH`.

**Security note:** This internal candidate has not been code-signed for production distribution. Do not turn off Windows security protections to force an installation. Record warnings or blocks during authorized testing.

### 3. Create a Markdown file

Save a file named `example.md` with content such as:

````markdown
# Quarterly Project Update

## Overview

This document was written in Markdown and converted with Finisova.

## Progress

| Workstream | Status |
|---|---|
| Development | Complete |
| Verification | In progress |

## Equation

Einstein's mass–energy relation is $E = mc^2$.

## Next steps

- Review the draft.
- Check tables and headings.
- Share the final Word document.
````

### 4. Convert using the GUI

1. Launch **Finisova** from its installed location or Start Menu shortcut.
2. Select a Markdown input file (or use another supported input workflow).
3. Choose the output format: **DOCX**, **PDF**, or **Both**.
4. Select an output location and, if needed, a supported presentation preset/profile/template.
5. Optionally inspect the preview, understanding that it is not Word's final layout.
6. Start the conversion and inspect the resulting artifact(s) and diagnostics.

For the first smoke test, choose **DOCX**. Use **PDF** or **Both** only on a machine where Microsoft Word is installed and correctly licensed.

## Command-line interface

Finisova includes a standalone console executable for scripting and repeatable conversions:

```powershell
$cli = Join-Path $env:LOCALAPPDATA "Programs\Finisova\cli\finisova-cli.exe"
& $cli --help
```

The examples below use the recorded CLI command shape. The **installed program's `--help` output is authoritative** for exact options and validation.

### Convert one Markdown file to DOCX

```powershell
& $cli ".\example.md" --format docx --no-open
```

### Convert to PDF (requires Microsoft Word)

```powershell
& $cli ".\example.md" --format pdf --no-open
```

### Generate both DOCX and PDF

```powershell
& $cli ".\example.md" --format both --no-open
```

### Specify an output directory

```powershell
& $cli ".\example.md" --format docx --output ".\exports" --no-open
```

### Process a directory of Markdown files

The accepted CLI includes a directory-input workflow for Markdown files:

```powershell
& $cli ".\markdown-documents" --format docx --no-open
```

Check `--help` and the resulting diagnostics for configuration, supported batch behavior, and output naming. CLI conversion errors and partially produced results should be handled based on the actual exit code and messages rather than assuming success.

**Notes:**

- `--no-open` suppresses automatic opening of the generated file.
- `--format` supports `docx`, `pdf`, and `both`.
- **There is no `--version` CLI option** in the accepted 2.1.0 CLI contract.
- Automatic `PATH` registration is not provided by the current installer.
- CLI style-preset selection is not part of the completed 2.1.0 CLI capability.

## How conversion works

The principal document pipeline is:

```text
Markdown source
       |
       v
Canonical parsing and document structure
       |
       v
Metadata + presentation resolution
       |
       v
Word-native DOCX rendering
       |
       +-------------------> DOCX output
       |
       +--> Microsoft Word native PDF export
                                      |
                                      v
                                   PDF output
```

The PDF route **consumes the DOCX result**. It is not an independent Markdown-to-PDF compiler. This design is intended to keep supported structure and formatting consistent across output choices, while retaining the limitations of Microsoft Word's PDF renderer.

## Privacy and local processing

Finisova is designed around **local desktop conversion**, not a cloud service requiring you to upload source documents.

In ordinary Markdown-to-DOCX use, your files are selected and processed on the local computer. Some preview and diagram components use bundled local assets. Explicitly opening an external HTTP/HTTPS link is a different action and may contact that website.

**Verification boundary:** This description does not claim that the full installed Build5 package has passed network-isolation, telemetry, security-penetration or all-offline runtime qualification. Those checks remain part of release readiness. Do not rely on this README alone as a security or regulatory assurance.

## Known limitations and verification status

| Area | Current limitation or qualification |
|---|---|
| General distribution | **Build5 is internal only.** Production release, signing and public distribution are not authorized. |
| Clean-machine install | Not yet verified in an independently provisioned Windows installation. |
| Word absent | The DOCX-without-Word design is accepted, but **Build5's installed Word-absent behavior remains to be tested**. |
| PDF backend | Requires Microsoft Word. LibreOffice PDF export is not included in 2.1.0. |
| CLI styling | GUI preset management is available; CLI style-preset selection remains deferred/incomplete. |
| Preview fidelity | Semantic rather than pixel-perfect DOCX/PDF rendering. Some supported content may display differently. |
| SVG into DOCX | The end-to-end SVG data-URI-to-DOCX embedding path was **NOT VERIFIED** in Build5. |
| Offline packaged preview | Installed, offline Math/Mermaid preview has not completed clean-machine acceptance. |
| Native/runtime behavior | Internal build/smoke checks passed; external environment and long-session stability need further QA. |
| Word COM post-processing | The Build5 default uses `word_com=False`; the optional Word-refresh/teardown path is **not** proven by the normal default-conversion smoke test. |
| Security/TLS | Build5 intentionally excludes certain OpenSSL libraries based on its packaging dependency analysis. Do not assume general-purpose TLS/HTTPS support for internal components without feature-specific verification. |
| Third-party compliance | §19A–§19E notices, source/relinking and redistribution obligations, including MSVC runtime review, remain open before distribution. |

The Build5 internal record includes 36/36 focused packaging tests, 42/42 bounded integration tests, and successful local GUI/CLI/installer build checks. **These are engineering-scope results, not an external end-to-end quality certification.**

## Troubleshooting

### PDF export says Microsoft Word is missing

DOCX generation can run without Word. For PDF or Both, install and configure a supported, licensed desktop copy of Microsoft Word on the authorized test computer. LibreOffice is not a substitute in Finisova 2.1.0.

### Finisova CLI is not recognized as a command

The current installer intentionally does not modify `PATH`. Run the executable with its full path:

```powershell
& "$env:LOCALAPPDATA\Programs\Finisova\cli\finisova-cli.exe" --help
```

If the file does not exist, record the installed path and setup log rather than copying executables from the developer build directory.

### Conversion produced a file but the result looks different

Inspect the generated document in Word. Check supported Markdown constructs, installed fonts, chosen template/profile/preset, document fields, and any diagnostics. Preview does not guarantee identical pagination.

### A diagram or SVG is missing

Review the exact Markdown syntax and conversion diagnostics. Mermaid preview/rendering supports a bounded subset; SVG embedding into DOCX is not fully verified for this Build5 candidate. Do not interpret a valid DOCX file as proof that every diagram or embedded image was rendered.

### The installer is blocked or reports a missing runtime

Record the exact Windows warning, error text, operating-system version and installation steps. Do not disable Windows Defender, SmartScreen, UAC or other security settings merely to make an internal build run.

### Output files cannot be created

Choose a writable output directory and verify that the output is not open or locked by another program. Finisova is designed to report unavailable output locations rather than silently pretending a file was produced.

## Licensing and distribution

**Finisova is commercial-oriented software; availability of source code or an internal installer does not imply permission to redistribute it.** Product use is governed by its applicable **End User License Agreement (`EULA.txt`)**. The package also contains **`THIRD_PARTY_NOTICES.txt`** and related third-party license texts.

- The packaged publisher label is **Arculoric**; this label alone does not establish the legal registration status of a company.
- Build5 has **not** received production code-signing, license-compliance clearance or public-distribution approval.
- No trial length, purchase price, activation method, renewal policy, or upgrade entitlement should be inferred from this README; those commercial terms have not been finalized here.
- The GitHub repository's legal license (if any) must be evaluated separately from the installer EULA and third-party notices.

**Do not upload this frozen internal installer to a public Release or a public Git repository until the authorized release and third-party compliance gates are satisfied.**

## Internal build identity

For authorized testers and auditors only:

| Item | Value |
|---|---|
| Product/build | Finisova 2.1.0 — Build5 |
| Candidate type | Frozen internal evaluation |
| Installer | `Finisova_v2.1.0_Setup.exe` |
| Installer size | 114,101,262 bytes |
| Installer SHA-256 | `2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d` |
| Build baseline | `c515974614d6323abadd5e14217b73a7387b3416` plus a specifically recorded hashed overlay |
| Build5 validation | GUI and CLI payloads, D-1/D-2/D-3, local smoke and installer compilation passed |
| Source reproducibility | Requires the exact recorded overlays, packaging revisions, runtime provenance and build environment; **not reproducible from Git HEAD alone** |
| Remaining gate | Clean Windows Installer Qualification and legal/third-party compliance |
| Production status | **Not released** |

The accepted Build5 snapshot and artifact manifests remain the source of truth for its detailed provenance. This table is informational and does not replace them.

## Roadmap and release readiness

The 2.1 release train focuses on **professional styling, Word-native PDF export, advanced preview, and Windows delivery verification**.

At the time of this README:

- The 2.1 implementation and internal packaging milestones have advanced through Build5.
- The frozen Build5 candidate is eligible for a **separately authorized** clean Windows installer test.
- Installed feature verification, complete third-party licensing review, RC freeze, final Human Acceptance, and production release are still separate gates.
- Planned or discussed future enhancements—such as alternative PDF engines, additional import/conversion directions, licensing/activation mechanisms, or broader platform compatibility—are **not** part of the confirmed 2.1.0 shipping contract.

## Feedback and issue reports

For authorized testers, a useful report includes:

1. Operating-system edition and version; Microsoft Word present/absent.
2. Finisova version and verified installer SHA-256.
3. Input type and a **sanitized**, minimal Markdown sample (never attach private documents without permission).
4. Selected output mode and style/template/profile.
5. Steps to reproduce, observed result, expected result, and exact error message.
6. Whether the issue occurs in GUI, CLI, preview, or installed/uninstalled behavior.

The project repository is [Quansheng2022/finisova](https://github.com/Quansheng2022/finisova). Do not assume that public issues, release assets, or support channels are active until explicitly configured.

---

**Finisova — professional Markdown-to-Word workflows, built for local use.**

*README prepared for version 2.1.0 based on the documented Build5 internal engineering state. Public-facing claims and distribution links require a separate release review.*
