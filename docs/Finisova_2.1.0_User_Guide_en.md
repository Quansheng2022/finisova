# Finisova 2.1.0 User Guide

**Windows Graphical User Interface and Command-Line Guide**  
**Document version:** 1.0 · **Prepared:** October 16, 2026  
**Applicable build:** Finisova 2.1.0 Build5 (internal evaluation candidate)  
**Publisher identifier:** Arculoric  
**Document status:** Internal evaluation user guide; installation procedures, user-interface details, and the supported environment must be revalidated before commercial release.

> **Important notice**  
> Build5 has passed GUI/CLI packaging, limited smoke tests, and packaging checks on the development machine. It has **not** completed independent clean-Windows installation/uninstallation qualification, separate Word-absent testing, comprehensive installed/offline preview verification, or third-party redistribution-license clearance. Build5 is **not a public production release**. This guide supports **separately authorized internal evaluation** only. Reading it does not grant permission to install, distribute, publish, or close the P24-03 release gate.

---

## Guide Contents

1. Introduction and Use Cases
2. Features and Supported Boundaries
3. System Requirements and Microsoft Word Dependencies
4. Obtaining, Verifying, and Installing the Internal Candidate
5. Your First Conversion: Markdown to Word in 10 Minutes
6. Using the Graphical Interface (GUI)
7. Using the Command-Line Interface (CLI) and Automation
8. Choosing DOCX, PDF, or Both and Interpreting Results
9. Style Presets, Profiles, Templates, and Project Mode
10. Equations, TOC, Captions, References, and Complex Documents
11. Advanced Preview, Mermaid, and Offline Resources
12. Files, Paths, Overwrite Protection, and Batch Jobs
13. Uninstallation, Reinstallation, and File Preservation
14. Privacy, Security, and Data Handling
15. Frequently Asked Questions and Troubleshooting
16. Verification Status, Known Limitations, and Build Identity
17. Glossary, Issue Template, and Test Checklist

---

## 1. Introduction and Use Cases

Finisova is a local Markdown document-conversion tool for Windows users. It turns `.md` drafts into editable `.docx` documents that can be reviewed, refined, and delivered in Microsoft Word. If a supported, licensed copy of Microsoft Word is installed and available, Finisova can also create PDF files using Word's native export capability.

Common use cases include:

- **Technical documentation:** Turn Markdown notes, system specifications, and operating procedures into Word deliverables.
- **AI-assisted drafts:** Convert Markdown drafted with AI assistance into structured Word files for human review.
- **Business reports:** Use headings, tables, a table of contents, and style presets to achieve consistent formatting.
- **Research and academic work:** Use supported editable equations, captions, cross-references, citations, and bibliographies.
- **Groups of documents:** Apply consistent document processing to multiple Markdown files or a project.

### 1.1 How the conversion works

```text
Markdown source
      |
      v
Structure parsing / metadata / presentation resolution
      |
      v
Native Word DOCX rendering and document post-processing
      |
      +-- DOCX output (Microsoft Word not required to create it)
      |
      +-- Microsoft Word native PDF export --> PDF output
```

PDF conversion follows a **DOCX-first** pipeline: Finisova does not use a separate Markdown-to-PDF semantic compiler. Fonts, pagination, and field evaluation in the resulting PDF are consequently dependent on the installed Microsoft Word environment.

### 1.2 What Finisova is not

Finisova is not a full Word editor, a cloud document-hosting service, an arbitrary HTML/CSS typesetting system, or a guarantee that every Markdown dialect, mathematical command, or Mermaid diagram can be converted without changes. The preview is not an exact simulation of Word pagination.

## 2. Features and Supported Boundaries

| Feature | Entry point | Description |
|---|---|---|
| Markdown to DOCX | GUI / CLI | Creating an editable Word document does not require Microsoft Word. |
| Markdown to PDF | GUI / CLI | Uses Microsoft Word's native PDF export; Word must be available. |
| DOCX + PDF (`Both`) | GUI / CLI | Requests two artifacts in the same workflow; partial success is possible. |
| File selection and supported pasted input | GUI | Use supported files or text inputs; confirm actual installed-screen labels. |
| Batch and directory conversion | CLI / supported GUI workflows | Processes Markdown files in a directory or defined job; check per-file results. |
| Profiles and Word templates | GUI / existing configuration | Provide reusable document presentation and formatting. |
| User-defined style presets | GUI | Select, edit, save, import/export, reset, and inherit supported presets. |
| CLI style-preset selection | — | **Not completed in 2.1.0**; do not assume `--style-config` is available. |
| TOC, metadata, editable equations | DOCX pipeline | Uses supported Markdown and configuration semantics. |
| Captions, cross-references, citations | DOCX pipeline | Bounded syntax; review results in the generated document. |
| Advanced Preview | GUI | Supports selected formulas, Mermaid, links, and styles; not Word-equivalent layout. |
| PDF without Word | — | No LibreOffice fallback in 2.1.0; missing Word must be diagnosed. |

**Build distinction:** These are implemented or accepted product capabilities. This does **not** mean that every feature has passed independent clean-machine verification of the installed Build5 package. See Chapter 16 for the actual verification status.

## 3. System Requirements and Microsoft Word Dependencies

### 3.1 Basic requirements

- **Operating system:** Windows desktop. Windows 11 is the current priority for independent clean-install QA; a final list of supported editions and releases has not yet been approved.
- **Free disk space:** Sufficient space for the installer, application files, temporary files, and output documents. A formal minimum has not been published.
- **Permissions:** Read access to Markdown inputs and write access to the selected output directory.
- **Python:** End users are not normally expected to install Python when using the packaged application; this remains subject to clean-machine verification.
- **Microsoft Word:** Not needed to *create* DOCX. Required for PDF output and the PDF part of `Both`; a correctly installed/licensed desktop edition must be available.

### 3.2 Word-dependency matrix

| Workflow | Without Word | With Word |
|---|---|---|
| DOCX | Supported by design; independent Build5 Word-absent test remains open | Available |
| PDF | Unavailable; an explicit Word-dependency error is expected | Uses Word's native PDF export |
| Both | DOCX may be produced while the PDF branch reports failure or dependency | Can produce both artifact types |
| Final Word TOC/field refresh | Recipient may need to update fields in Word | Fields can be inspected and refreshed in Word |

**Important:** The LibreOffice PDF backend was deferred beyond 2.1.0. Installing LibreOffice does not replace the Microsoft Word requirement.

## 4. Obtaining, Verifying, and Installing the Internal Candidate

### 4.1 Identify the correct installer

The current internal candidate installer is named:

```text
Finisova_v2.1.0_Setup.exe
```

Recorded Build5 binary identity:

| Item | Recorded value |
|---|---|
| Installer size | 114,101,262 bytes (approximately 108.8 MiB) |
| SHA-256 | `2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d` |
| Status | FROZEN — INTERNAL CANDIDATE ONLY |

Open PowerShell in the folder containing the installer, then run:

```powershell
Get-FileHash ".\Finisova_v2.1.0_Setup.exe" -Algorithm SHA256
```

If the value differs, **stop**. Do not rename a different installer to make it look like Build5.

### 4.2 If the installer is supplied as split `.7z` volumes

Place every `.7z.001`, `.7z.002`, etc. part in the same folder. Extract the first volume with 7-Zip, allowing the extractor to use the others automatically.

```text
Finisova_v2.1.0_Setup.7z.001
Finisova_v2.1.0_Setup.7z.002
...
```

Do not attempt to extract only `.002` or omit the final part. The number of volumes may vary depending on compression settings and volume size. **Verify the recovered `.exe` SHA-256 after extraction**; a count of files alone is not an integrity check.

### 4.3 Install only on an authorized test computer

1. Confirm that the test computer and installation/uninstallation actions have been explicitly authorized. Do not assume a production workstation is a disposable QA machine.
2. Verify the installer hash.
3. Start the installer. Review the EULA, displayed publisher information, and any operating-system security warnings.
4. Record the installer version, installation scope, shortcuts, and all unusual messages.
5. Complete the guided installation, then start Finisova from the Start Menu or the installed folder.

The recorded Inno Setup configuration specifies this **expected per-user** location:

```text
%LOCALAPPDATA%\Programs\Finisova
```

Expected executable locations:

```text
%LOCALAPPDATA%\Programs\Finisova\Finisova.exe
%LOCALAPPDATA%\Programs\Finisova\cli\finisova-cli.exe
```

The installer deliberately does **not** modify your `PATH` to include the CLI. Do not disable Microsoft Defender, SmartScreen, or UAC to make an internal candidate install; record warnings and send them to the QA owner.

### 4.4 First post-install checks

Confirm that the GUI starts, identifies itself as Finisova, and exits normally. Then run:

```powershell
$cli = Join-Path $env:LOCALAPPDATA "Programs\Finisova\cli\finisova-cli.exe"
& $cli --help
```

**The accepted CLI does not provide a `--version` option.** A failed `--version` invocation is not evidence of a broken installation.

## 5. Your First Conversion: Markdown to Word in 10 Minutes

### 5.1 Create a sample Markdown file

Using VS Code, Typora, Notepad, or another editor, save the following text as a UTF-8 file named `example.md`:

````markdown
# Quarterly Project Update

## Overview
This document was prepared in Markdown and converted with Finisova.

## Progress
| Area | Status |
|---|---|
| Development | Complete |
| Testing | In progress |

## Equation
Einstein's relation is $E = mc^2$.

## Next steps
- Review the content.
- Check the Word layout.
- Share the approved document.
````

### 5.2 Convert in the GUI

1. Start Finisova.
2. Select `example.md`, or use the supported drag-and-drop input.
3. Choose **DOCX** as the output format.
4. Select or confirm the output folder.
5. Keep the default profile and style for this first test.
6. Start the conversion using the control visible in your installed GUI and wait for completion.
7. Find the generated DOCX through the result shown in the GUI. Open it and check headings, the table, equation, and paragraphs.

**UI qualification:** Final installed-GUI screenshots and exact control labels have not yet been reconciled with this manual. The procedures deliberately describe functions rather than inventing button captions.

### 5.3 Convert in the CLI

On an authorized test computer where Finisova is installed:

```powershell
$cli = Join-Path $env:LOCALAPPDATA "Programs\Finisova\cli\finisova-cli.exe"
& $cli ".\example.md" --format docx --no-open
```

`--no-open` prevents automatic opening of the generated document by its Windows file association, which is useful for unattended testing. Check the actual result messages for output naming, destination directory, and diagnostics.

### 5.4 Checklist for your first result

- [ ] The application reports success or provides explicit diagnostic information.
- [ ] The output file exists at the reported location and has a reasonable file size.
- [ ] Microsoft Word can open the generated DOCX.
- [ ] Headings, the table, and body text are present as expected.
- [ ] There is no unexplained background process, locked document, or abnormal application termination.

## 6. Using the Graphical Interface (GUI)

### 6.1 Supplying content

The GUI supports Markdown file selection and retains supported drag-and-drop, pasted-content, project, and batch workflows. Inputs must match the capabilities of the actual control and workflow. Do not assume a single-file drop zone will accept arbitrary PDF/DOCX inputs, URLs, directories, or multiple files.

For AI-generated Markdown pasted into Finisova, first check code fences, heading levels, mathematical delimiters, and potentially unsupported HTML. When in doubt, save a clean UTF-8 `.md` file and convert it.

### 6.2 Output format and status

- **DOCX:** Produces an editable Word document and is the preferred first test of the conversion pipeline.
- **PDF:** Requires Microsoft Word and uses its native export route.
- **Both:** Requests both artifact types. If one fails, inspect each artifact's result instead of treating the presence of one file as complete success.

GUI states include converting, success, failure, and partial success. A warning may accompany an otherwise valid artifact. Interpret the specific diagnostic and artifact status rather than relying on a single general status label.

### 6.3 Output location

Choose a writable, easy-to-find directory, for example:

```text
Documents\Finisova\Exports
```

Avoid system-protected folders. Do not assume it is safe to overwrite a file that is currently open in Word. For filename collisions, automatic naming, relative paths, and inaccessible folders, use the actual installed application's feedback.

### 6.4 Styles and presets

The GUI's Style Preset workflow supports selecting, editing, saving, importing, exporting, resetting, and inheriting supported properties. Save a known-good preset under an identifiable name, then verify changes using a small Markdown file before adjusting many font and paragraph settings together.

**A preset is not a complete Word template.** Style Preset, Profile, Word Template, and Project Mode have specific precedence and supported-property rules. Unspecified properties inherit values determined by the existing presentation resolver.

### 6.5 Preview

Preview helps inspect Markdown structure and supported formulas, Mermaid diagrams, links, tables, and styles. Word pagination, headers and footers, font metrics, and complex fields may still look different. Always inspect the generated DOCX or PDF before delivery.

## 7. Using the Command-Line Interface (CLI) and Automation

### 7.1 General syntax

```powershell
& $cli <MarkdownFileOrDirectory> --format docx|pdf|both [other-supported-options]
```

This is *illustrative syntax*, not a command to execute verbatim. Replace placeholders with real paths. Obtain authoritative option details using:

```powershell
& $cli --help
```

### 7.2 Common commands

**Single DOCX:**

```powershell
& $cli ".\report.md" --format docx --no-open
```

**PDF (requires Word):**

```powershell
& $cli ".\report.md" --format pdf --no-open
```

**Both DOCX and PDF:**

```powershell
& $cli ".\report.md" --format both --no-open
```

**Specify an output directory:**

```powershell
& $cli ".\report.md" --format docx --output ".\exports" --no-open
```

**Convert a directory:**

```powershell
& $cli ".\markdown-documents" --format docx --no-open
```

**View help:**

```powershell
& $cli --help
```

### 7.3 Common CLI options

| Option | Meaning and considerations |
|---|---|
| `--format docx` | Produces DOCX; the legacy behavior without an explicit format remains DOCX by design. |
| `--format pdf` | Creates PDF from the DOCX pipeline using Microsoft Word. |
| `--format both` | Requests DOCX and PDF; check artifact-specific and partial-success results. |
| `--output` / `-o` | Selects output location; see actual `--help` and result messages for behavior. |
| `--no-open` | Does not automatically open produced files; suitable for testing and scripts. |
| `--config` / `-c` | Reads an accepted existing configuration; refer to the configuration contract for its structure. |
| `--verbose` / `-v` | Requests more detailed diagnostics. |

Other ASCII/preview options are advanced and should be used only as documented by the installed `--help`. CLI `--style-config` did **not** complete 2.1.0 acceptance; **do not depend on it in automation**.

### 7.4 Automation example: capture logs and exit status

```powershell
$cli = Join-Path $env:LOCALAPPDATA "Programs\Finisova\cli\finisova-cli.exe"
$src = Join-Path $PWD "report.md"
$dest = Join-Path $PWD "exports"

& $cli $src --format docx --output $dest --no-open 2>&1 |
    Tee-Object -FilePath ".\finisova-run.log"

$exitCode = $LASTEXITCODE
Write-Host "Finisova exit code: $exitCode"
if ($exitCode -ne 0) {
    Write-Warning "Review the diagnostics; do not assume all artifacts succeeded."
}
```

Before running, verify source paths, output permissions, available disk space, and whether logs may contain sensitive content. The recorded CLI exit-code contract is `0` for success, `1` for failure, and `3` for partial success; the installed executable and its accepted contract remain authoritative.

### 7.5 Avoid these scripting assumptions

- Do not rely on CLI `--version` or unfinished style-preset-selection options.
- Do not treat the mere presence of a file as proof of complete and correct conversion.
- Do not assume PDF is a Word-independent rendering engine.
- Do not automatically retry conversions by overwriting existing user files.
- Never terminate an existing user-owned `WINWORD.EXE` process as part of unattended cleanup.

## 8. Choosing DOCX, PDF, or Both and Interpreting Results

### 8.1 Which format should you choose?

| Goal | Recommended format | Why |
|---|---|---|
| Edit text, equations, and styles | DOCX | Remains editable in Microsoft Word. |
| Share a fixed-layout reading copy | PDF | Native Word export generally follows Word layout. |
| Keep both an editable and a reading copy | Both | Produces two artifact types in one workflow. |
| Convert on a machine without Word | DOCX | Do not expect PDF export to succeed without Word. |

### 8.2 Partial success in Both mode

If DOCX succeeds but Word is absent, Word fails, or PDF export fails, examine the GUI/CLI status and artifact-specific diagnostics to learn what was actually produced. Do not extrapolate DOCX success into PDF success.

### 8.3 Word table of contents and fields

Some Word-native features, such as TOC fields, cross-reference fields, or pagination-dependent details, may require updating inside Word. If page numbers or field results are stale, first save a copy of the document, then update the relevant fields in Word and inspect them.

Build5 uses `word_com=False` by default; an ordinary DOCX conversion therefore does not launch Word just to update the TOC. This does **not** prove that the optional Word COM field-refresh/cleanup path has been fully qualified.

## 9. Style Presets, Profiles, Templates, and Project Mode

### 9.1 Understand the four mechanisms

- **Profile:** A ready-made set of document presentation settings.
- **Word Template:** A supported `.docx` template or style resource used by the output document.
- **Style Preset:** Explicit overrides for a bounded set of properties such as headings, fonts, and paragraphs.
- **Project Mode:** Organizes project-level settings and multiple documents into a repeatable workflow.

These are **not** governed by a universal “last selection wins” rule. Finisova applies the existing Presentation Resolution contract to supported properties. Verify the actual output on a representative sample.

### 9.2 A reliable styling workflow

1. Generate a minimal DOCX with default styles.
2. Select an appropriate Profile and inspect headings and body text.
3. If a corporate layout is required, test the designated Word Template.
4. Apply Style Preset overrides only to properties you actually need to change.
5. Save or export the reusable preset and repeat the check with a representative longer document.

### 9.3 Project Mode and Batch

For a project containing multiple Markdown files, begin with a small sample. Verify file discovery, project configuration, image-relative paths, output locations, name collisions, failures, and partial success. The exact project manifest format and GUI menu paths must be taken from the accepted Project Mode specification and installed product; this guide does not invent an unverified manifest syntax.

## 10. Equations, TOC, Captions, References, and Complex Documents

### 10.1 Headings, lists, and tables

Use a consistent hierarchy of standard Markdown headings:

````markdown
# Report Title
## Background
### Detail

- Item A
- Item B

| Item | Amount |
|---|---:|
| A | 10 |
| B | 20 |
````

Keeping source structures clean improves the Word output. Complex merged cells, arbitrary embedded HTML, and unrestricted CSS are not guaranteed to reproduce identically in Word.

### 10.2 Inline and block equations

The accepted product implements a supported subset of `$...$` and `$$...$$` syntax. Supported expressions are intended to become native, editable Word equations in DOCX, rather than flat screenshots.

````markdown
The equation is $E = mc^2$.

$$
\frac{a+b}{c}
$$
````

Unsupported TeX commands, macros, or extensions may result in warnings or fallback rendering. Check the resulting equation's editability and layout in Word.

### 10.3 Table of contents (TOC)

Use consistent Markdown heading levels for long documents. Finisova's Word TOC is based on document structure. TOC depth, template settings, field refresh, and page-number calculation can affect its final appearance. Before delivery, inspect the TOC in Word and update fields if needed.

### 10.4 Captions and cross-references

Figure, table, and equation captions and their cross-references must follow the syntax supported by Finisova's canonical reference contract. Do not assume that an image's Markdown alt text, any arbitrary anchor, or ordinary bracketed text automatically becomes a numbered Word caption or cross-reference. Start with a previously validated example and check that each generated reference points to the intended target.

### 10.5 Citations and bibliography

Supported citation semantics include key-based expressions such as `[@key; @key]`. Bibliography materialization depends on the configured citation data, the supported subset of BibTeX, and the conversion workflow. Do not assume every `.bib` file is discovered automatically. Resolve duplicate or missing keys and BibTeX parse errors before conversion.

## 11. Advanced Preview, Mermaid, and Offline Resources

### 11.1 What does Preview support?

The accepted preview scope includes supported equations, Mermaid diagrams, links, headings, tables, code blocks, other rich content, and style-aware display. The preview uses the canonical document semantics and should display bounded fallbacks or diagnostics for unsupported input.

### 11.2 Mermaid example

````markdown
```mermaid
flowchart LR
    Draft[Markdown] --> Docx[DOCX]
    Docx --> Pdf[PDF via Word]
```
````

This example illustrates Markdown/Mermaid syntax. Supported diagram categories, node labels, and runtime rendering are limited to the implementation's accepted subset. In particular, **Build5's fully installed, offline preview has not yet completed independent Windows QA**.

### 11.3 Why Preview and PDF may differ visually

- Preview is **not** the Microsoft Word layout engine.
- Available fonts, page margins, and Word pagination affect final output.
- Word fields (such as TOC entries) may need refreshing in Word.
- The end-to-end SVG data-URI-to-DOCX embedding path was **not verified** for Build5.

### 11.4 Links and network access

Preview is designed with restrictions on link/resource access. Clicking an external `https://` link may deliberately access the Internet. “Local conversion” should not be interpreted as a promise that every possible user action is network-free. Avoid opening unfamiliar external links when working with sensitive material.

## 12. Files, Paths, Overwrite Protection, and Batch Jobs

### 12.1 Choose an output directory

Use a writable folder you can easily find and back up. Long paths, protected folders, read-only volumes, and files currently open in Word may cause failures. Preserve the original source and follow the reported diagnostics to choose a valid output location.

### 12.2 Relative image paths

Markdown image resolution may depend on the Markdown file's location, the project root, and the accepted image syntax. A simple layout is:

```text
report/
  report.md
  assets/
    chart.png
```

Reference the image with:

```markdown
![Chart](assets/chart.png)
```

After conversion, confirm that the image appears in the DOCX. The presence of a valid DOCX does not prove every image was successfully embedded.

### 12.3 Before running a batch conversion

- [ ] Back up the original Markdown sources and related assets.
- [ ] Run a representative small-sample conversion first.
- [ ] Confirm relative image, BibTeX, and project-configuration paths.
- [ ] Avoid collisions between identically named outputs.
- [ ] Check the available disk space.
- [ ] Read diagnostics for each document; do not equate partial success with complete success.

## 13. Uninstallation, Reinstallation, and File Preservation

**Only for explicitly authorized standalone QA computers.** Build5 clean-Windows installation, uninstallation, and reinstallation behavior has not received final qualification.

On an authorized test computer, use **Windows Settings → Apps → Installed apps** to locate Finisova and initiate the normal uninstall workflow. Record whether the app-owned directory, shortcuts, and uninstall entry are removed. Verify that user-created Markdown, Word, and PDF files remain safe.

Do not recursively clear `%LOCALAPPDATA%`, `Documents`, or other user directories to “confirm” uninstallation. Do not manually delete files of uncertain ownership. Before reinstallation, record the environment and installer SHA-256 for reproducibility.

## 14. Privacy, Security, and Data Handling

Finisova is intended to process documents locally rather than requiring upload of Markdown content to an external cloud conversion service. Individual functions may still interact with the operating system, Microsoft Word, local rendering components, and user-initiated external web links.

This guide **does not make unverified promises** of zero network traffic, absence of telemetry, completed penetration testing, absolutely no residual data, universal platform compatibility, or compliance with a particular regulatory framework. Those claims require independent release-stage validation.

During internal testing, prefer synthetic sample documents. Do not attach client secrets, passwords, personally identifiable information, confidential business information, or unauthorized files to issue reports.

## 15. Frequently Asked Questions and Troubleshooting

| Symptom | Possible cause | Recommended action |
|---|---|---|
| PowerShell cannot find `finisova-cli.exe` | Installer does not add CLI to PATH | Invoke the full CLI path shown in this guide. |
| PDF export reports missing Word | Word is not available or configured correctly | Use DOCX first; verify Word on an authorized test computer. |
| DOCX succeeds but PDF fails | PDF branch of a Both job failed | Read partial-result diagnostics; do not report complete success. |
| Output file cannot be found | Output directory or generated filename differs from expectation | Use the path reported by the conversion result. |
| Cannot write output / file is locked | Missing permission or file open in Word | Choose a writable location and avoid overwriting an open file. |
| Formula or Mermaid Preview looks different from Word | Rendering capabilities and layout engines differ | Inspect final DOCX/PDF as the delivery artifact. |
| TOC page numbers seem incorrect | Word fields have not been refreshed | Save a copy, then update the appropriate fields in Word. |
| SVG or diagram is missing | Unsupported or unverified syntax/path | Supply a minimal reproducible Markdown sample; do not assume universal SVG support. |
| Windows Defender / SmartScreen warns during install | Unsigned internal candidate, reputation, or security policy | Preserve safeguards and record the warning for QA. |
| CLI `--version` is not recognized | Not part of the accepted CLI interface | Use `--help` and verify the installer hash for identity. |
| Preview or exit causes a crash | GUI/worker environment or native-component issue | Record exit code, OS version, reproduction steps, and logs. |

### 15.1 What should an issue report contain?

1. Windows edition/version and whether Microsoft Word is installed (include its version if available).
2. Installer SHA-256, installed path, and offline/online test conditions.
3. Exact steps for GUI, CLI, Preview, Install, or Uninstall.
4. A **sanitized minimal Markdown reproducer** that can safely be shared.
5. Selected format, Profile, Template, and Style Preset.
6. Actual error, exit code, diagnostic text, and artifact validity.

## 16. Verification Status, Known Limitations, and Build Identity

### 16.1 Recorded Build5 engineering checks

| Evidence area | Result |
|---|---|
| GUI / CLI PyInstaller builds | PASS |
| Inno Setup installer build | PASS |
| Qt component D-1, license metadata D-2, native provenance D-3 | PASS |
| Controlled Cairo runtime | 19/19 DLLs matched expected hashes |
| GUI launch/close, CLI help, CLI DOCX smoke test | PASS (developer machine) |
| Packaging focused tests | 36/36 PASS |
| Bounded integration tests | 42/42 PASS |
| New failures within executed acceptance scope | 0 |
| Post-build audit | ACCEPTED WITH DISCLOSED QUALIFICATIONS |

### 16.2 Outstanding QA and compliance tasks

| Item | Status |
|---|---|
| Independent Windows clean install / uninstall / reinstall | NOT VERIFIED |
| Separate computer without Word: DOCX and PDF dependency behavior | NOT VERIFIED |
| Offline installed Math/Mermaid Preview | NOT VERIFIED |
| SVG data URI → DOCX embedded image | NOT VERIFIED |
| Optional Word COM field refresh and complete process-ownership cleanup | Not covered by the default smoke test |
| Commercial third-party license obligations §19A–§19E | OPEN |
| P24-03 parent closure / P24-04 RC freeze / production release | NOT AUTHORIZED |

**Legal and distribution note:** Bundled EULA and third-party license texts do not, on their own, establish that all redistribution obligations have been met. Trial length, license-key activation, pricing, upgrade rights, and public download arrangements are not established in this guide.

### 16.3 Internal candidate identity

```text
Product        : Finisova 2.1.0
Candidate      : Build5 (internal only)
Installer name : Finisova_v2.1.0_Setup.exe
Installer size : 114101262 bytes
SHA-256        : 2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d
Build baseline : c515974614d6323abadd5e14217b73a7387b3416 + hashed overlay revisions
```

**Reproducibility boundary:** Build5 consists of the specified Git HEAD, separately hashed packaging/source overlays, and controlled native runtime components. **The same binary cannot be assumed to be reproducible from that Git HEAD alone** in a new checkout.

## 17. Glossary, Issue Template, and Test Checklist

### 17.1 Glossary

| Term | Meaning |
|---|---|
| DOCX | An editable document format used by Microsoft Word. |
| PDF | A fixed-layout format primarily used for reading and distribution. |
| Both | A single request to generate DOCX and PDF artifacts. |
| Profile | A predefined set of document-presentation settings. |
| Style Preset | Reusable formatting overrides for supported properties. |
| Word Template | A Word document template or supported style resource. |
| TOC | Table of Contents. |
| Caption | A labeled/numbered figure, table, or equation description. |
| Cross-reference | A reference to another document heading or object. |
| Citation/Bibliography | In-text citations and their bibliography entries. |
| Artifact | An actual output file, such as a DOCX or PDF, produced by a conversion request. |
| Internal Candidate | A build reserved for authorized internal verification; not a public release. |

### 17.2 Minimal issue-report template

```text
Finisova version/build:
Installer SHA-256:
Windows version:
Microsoft Word installed? (Y/N):
Workflow (GUI / CLI / Preview / Install / Uninstall):
Source Markdown (sanitized sample):
Selected output format:
Observed result / exit code:
Expected result:
Artifacts produced (paths and validity):
Screenshots / relevant logs (redacted):
```

### 17.3 Internal installed-Windows QA checklist

- [ ] Test machine is explicitly authorized and contains no at-risk production/user data.
- [ ] The original installer hash has been verified.
- [ ] Installed location, shortcuts, and publisher/version metadata have been recorded.
- [ ] GUI and CLI run from **installed paths**, not the source checkout.
- [ ] DOCX conversion and document readability have been checked.
- [ ] Word-present PDF/Both workflows have been tested independently where required.
- [ ] Word-absent behavior has been tested on a separate environment, **not** by uninstalling the user's Word.
- [ ] Offline installed Preview has been checked in the required scenarios.
- [ ] Uninstallation and reinstallation preserve user-created files.
- [ ] Outstanding SVG, §19 licensing, and other known limitations are documented.

---

**Finisova — Convert Markdown into polished, editable Word documents.**  
**Document basis:** Accepted Build5 specifications and audit (October 16, 2026). Verify final UI labels, compatibility, and license terms before release.
