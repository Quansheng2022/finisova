# Finisova 2.1.0 使用手册

**Windows 图形界面与命令行使用指南**  
**文档版本：** 1.0 · **编写日期：** 2026-10-16  
**适用构建：** Finisova 2.1.0 Build5（内部测试候选）  
**发布者标识：** Arculoric  
**文档状态：** 内部测试版使用手册；正式商业发布前须重新核对安装步骤、界面和支持范围

> **重要声明**  
> Build5 已完成开发机上的 GUI/CLI 构建、有限 Smoke Test 和 Packaging 验证，但尚未完成独立 Windows 电脑的安装与卸载资格验证、Word-absent 验证、完整离线预览验证及第三方许可合规审查。Build5 **不是正式公开发行版**。本手册可用于准备和执行**另行获得授权**的内部测试；阅读本手册本身不构成安装、分发、发布或关闭 P24-03 的授权。

---

## 阅读导航

1. 产品简介与适用场景
2. 功能一览和已知边界
3. 环境要求与 Microsoft Word 依赖
4. 获取、校验与安装（授权测试环境）
5. 10 分钟上手：Markdown → Word
6. 图形界面（GUI）操作
7. 命令行（CLI）操作与自动化
8. DOCX、PDF、Both 的选择和结果判读
9. 样式预设、Profiles、Templates 与 Project Mode
10. 公式、目录、标题、引用及其他复杂文档
11. 高级预览、Mermaid 与离线资源
12. 文件、路径、覆盖保护与批量处理
13. 卸载、重装与文件保留
14. 隐私、安全与数据处理边界
15. 常见问题与故障排查
16. 验证状态、已知限制与版本身份
17. 术语表、反馈模板与测试核对表

---

## 1. 产品简介与适用场景

Finisova 是面向 Windows 用户的本地 Markdown 文档转换工具，帮助你将 `.md` 草稿整理为可在 Microsoft Word 中继续编辑、审阅和交付的 `.docx` 文件。安装并配置 Microsoft Word 后，还可通过 Word 的原生导出功能输出 PDF。

典型使用场景：

- **技术文档**：将 Markdown 笔记、系统说明、操作手册整理成 Word 交付件。
- **AI 生成草稿**：把 AI 辅助生成的 Markdown 转为结构化 Word 文档，再由人工审阅。
- **工作报告**：通过标题、表格、目录和样式预设提高排版一致性。
- **研究与学术材料**：在受支持的语法范围内使用可编辑公式、图表标题、交叉引用、引用和参考文献。
- **成批文档**：对一组 Markdown 文件或项目中的内容进行一致化处理。

### 1.1 工作流程

```text
Markdown 原稿
    ↓
结构解析 / 元数据解析 / Presentation Resolution
    ↓
Word 原生 DOCX 渲染与文档后处理
    ├── DOCX 文件（不要求本机安装 Word）
    └── Microsoft Word 原生 PDF 导出 → PDF 文件
```

Finisova 的 PDF 功能采用 **DOCX-first** 路线，不使用一套独立的 Markdown-to-PDF 语义解析器。由于 Word 的字体、分页和字段计算机制，PDF 版式仍依赖实际的 Word 环境。

### 1.2 不属于本产品的功能

Finisova 不是完整的 Word 编辑器、在线文档托管平台、任意 HTML/CSS 排版器，也不保证将所有 Markdown 方言、公式和 Mermaid 图表无损转换。预览不是 Word 的精确分页模拟器。

## 2. 功能一览和已知边界

| 功能 | 使用入口 | 说明 |
|---|---|---|
| Markdown → DOCX | GUI / CLI | Word 文档创建不要求安装 Microsoft Word。 |
| Markdown → PDF | GUI / CLI | 使用 Microsoft Word 原生 PDF 导出；需要可用的 Word。 |
| DOCX + PDF（Both） | GUI / CLI | 同一转换工作流生成两类输出；可能有部分成功。 |
| Markdown 文件选择 / 粘贴 | GUI | 文件或其他已支持的输入方式；控件名称以实际安装版为准。 |
| 批量与目录转换 | CLI / 受支持的 GUI 工作流 | 可以处理目录内的 `.md` 文件；留意结果清单。 |
| Profiles、Word Templates | GUI / 既有配置 | 决定专业 Word 输出的字体、页面和文档表现。 |
| User-Defined Style Presets | GUI | 选择、编辑、保存、导入/导出、重置/继承受支持的预设。 |
| CLI Style Preset 选择 | — | **2.1.0 尚未完成**；不能假定存在 `--style-config`。 |
| TOC、Metadata、可编辑公式 | DOCX 转换链路 | 使用受支持的 Markdown 及配置语义。 |
| Captions / Cross References / Citations | DOCX 转换链路 | 有界语法，完整准确性需根据真实文档检查。 |
| Advanced Preview | GUI | 支持的公式、Mermaid、链接及样式；不等于 Word 排版。 |
| 无 Word 环境 PDF | — | 不提供 LibreOffice 替代后端；应明确报告 Word 依赖。 |

**版本区别：** 上表描述已经实现或纳入接受范围的产品能力，并不等于这些能力已经在 Build5 的全新安装机器上全部验证通过。有关具体测试结果见第 16 章。

## 3. 环境要求与 Microsoft Word 依赖

### 3.1 必要条件

- **操作系统：** Windows 桌面环境；Windows 11 是当前独立安装测试的优先目标。具体支持版本清单待正式 QA 冻结。
- **磁盘空间：** 需要为安装程序、程序文件和转换输出预留空间；正式最小磁盘容量尚未公布。
- **文件权限：** 必须能够读取 Markdown 文件，并能写入所选输出文件夹。
- **Python：** 已打包的应用通常不要求使用者自行安装 Python；最终仍须用干净 Windows 验证。
- **Microsoft Word：** 创建 DOCX 不依赖 Word；导出 PDF 或 Both 中的 PDF 则要求可用、授权的 Word 桌面程序。

### 3.2 功能与 Word 关系

| 场景 | 无 Word | 有 Word |
|---|---|---|
| DOCX | 设计上可用；Build5 独立无 Word 机器待测 | 可用 |
| PDF | 不支持；应提示缺少 Word | 可通过 Word 原生导出 |
| Both | DOCX 可生成，PDF 分支应报告依赖或失败 | 可输出两种文件 |
| Word 目录字段最终刷新 | 可能需由收件人在 Word 中更新 | 可在 Word 中检查、更新字段 |

**注意：** 在 2.1.0 中 LibreOffice PDF 后端已经延期，因此安装 LibreOffice 并不能替代 Microsoft Word。

## 4. 获取、校验与安装（授权测试环境）

### 4.1 确认安装包身份

当前内部候选安装程序名称为：

```text
Finisova_v2.1.0_Setup.exe
```

Build5 已记录的文件大小和 SHA-256：

| 项目 | 值 |
|---|---|
| 大小 | 114,101,262 bytes（约 108.8 MiB） |
| SHA-256 | `2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d` |
| 状态 | FROZEN — INTERNAL CANDIDATE ONLY |

在 PowerShell 中，先进入 Installer 所在目录，再运行：

```powershell
Get-FileHash ".\Finisova_v2.1.0_Setup.exe" -Algorithm SHA256
```

如果哈希值与记录不一致，停止使用该文件，不要尝试通过重新命名来“修复”。

### 4.2 如果安装包是多个 `.7z` 分卷

将 `.7z.001`、`.7z.002` 等所有分卷放在同一个文件夹，使用 7-Zip 对第一个分卷执行 Extract。不要仅解压 `.002`，也不要漏掉最后一个分卷。恢复 EXE 后**再次执行 SHA-256 校验**。

```text
Finisova_v2.1.0_Setup.7z.001
Finisova_v2.1.0_Setup.7z.002
...
```

分卷的数量可能因压缩设置和卷大小不同而变化；不要仅凭文件数量判断其完整性。

### 4.3 在授权测试电脑上安装

1. 先确认测试电脑及安装/卸载操作已经取得明确授权。不要默认在日常工作电脑上试验。
2. 验证 Installer 的 SHA-256。
3. 启动安装程序，阅读 EULA、发布者信息和操作系统安全提示。
4. 记录显示的版本、安装范围、快捷方式以及任何异常提示。
5. 按安装界面完成步骤，随后从开始菜单或安装目录启动 Finisova。

当前 Inno Setup 配置的**预期** per-user 安装目录：

```text
%LOCALAPPDATA%\Programs\Finisova
```

预期可执行文件：

```text
%LOCALAPPDATA%\Programs\Finisova\Finisova.exe
%LOCALAPPDATA%\Programs\Finisova\cli\finisova-cli.exe
```

Installer 不会自动把 CLI 路径加入系统 PATH，也不应因此要求你手工关闭 Defender、SmartScreen 或 UAC。遇到安全警告请记录其具体文本，交由测试负责人判断。

### 4.4 安装后初步检查

建议先验证 GUI 是否启动、窗口标题是否显示 Finisova，以及能否正常关闭。然后在 PowerShell 中执行：

```powershell
$cli = Join-Path $env:LOCALAPPDATA "Programs\Finisova\cli\finisova-cli.exe"
& $cli --help
```

**CLI 没有已接受的 `--version` 选项。** 不要把 `--version` 失败误判为安装损坏。

## 5. 10 分钟上手：Markdown → Word

### 5.1 创建示例 Markdown

使用 VS Code、Typora、Notepad 或其他文本编辑器，保存为 `example.md`（UTF-8）：

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

### 5.2 使用 GUI 转换

1. 打开 Finisova。
2. 选择 `example.md`，或使用界面支持的拖放输入。
3. 输出格式选择 **DOCX**。
4. 选择或检查输出文件夹。
5. 保留默认 Profile/Style，进行第一次测试。
6. 点击界面实际显示的转换操作，并等待任务结束。
7. 根据 GUI 的结果状态打开生成的 DOCX，检查标题、表格、公式及段落。

**注意：** GUI 的实际按钮名称和控件布局尚未经过最终安装版截图逐项核验。本手册采用功能名称描述操作，不凭空指定未经证实的按钮文案。

### 5.3 使用 CLI 转换

在已经完成授权安装的 Windows 电脑上：

```powershell
$cli = Join-Path $env:LOCALAPPDATA "Programs\Finisova\cli\finisova-cli.exe"
& $cli ".\example.md" --format docx --no-open
```

`--no-open` 会阻止自动通过文件关联打开输出文件，便于无人值守测试。默认输出目录、自动命名及失败诊断以实际程序结果为准。

### 5.4 第一次输出的检查清单

- [ ] 程序显示转换成功或明确的诊断信息。
- [ ] 输出路径存在且文件大小合理。
- [ ] 生成的文件可以在 Word 中打开。
- [ ] 标题层级、表格与正文没有明显丢失。
- [ ] 未出现非预期的程序残留、文档锁定或异常退出。

## 6. 图形界面（GUI）操作

### 6.1 输入内容

GUI 支持单个 Markdown 文件选择，并延续文件拖放、受支持的粘贴和项目/批处理工作流。输入必须是产品所支持的文件或文本形式。不要把任意 PDF、DOCX、目录、URL 或多个文件拖入单文件选择区域后假定它们都能被接受。

对于粘贴的 AI Markdown，建议先检查代码围栏、标题层级、数学分隔符以及是否包含不受支持的 HTML；必要时先保存为 UTF-8 `.md` 再转换。

### 6.2 输出格式与状态

- **DOCX：** 只生成 Word 文档，是验证转换主链路的首选。
- **PDF：** 需要 Microsoft Word；使用原生导出路径。
- **Both：** 同一请求包含两种输出。若其中一项失败，应阅读具体的 Artifact 状态，不可仅凭其中一个文件存在就判定全部成功。

GUI 会呈现转换中、成功、失败及部分成功等状态。某些诊断可以是 Warning，而不是整个请求失败；需根据具体输出结果判断。

### 6.3 输出位置

建议使用有写权限、容易定位的文件夹，例如：

```text
Documents\Finisova\Exports
```

不要选择系统保护目录，也不要在 Word 正打开同名输出文件时覆盖或重复生成。对输出文件名、重名处理、相对路径和未写入目录的失败行为，以实际 GUI 反馈为准。

### 6.4 风格与预设

在 GUI 提供的 Style Preset 入口，可以选择、编辑、保存、导入、导出与重置/继承受支持的样式。建议将初始设置另存为明确名称，修改后用小样本转换验证，不要一次改变大量字体和段落属性。

**样式并不等于完整 Word 模板。** Style Preset、Profile、Template 与 Project Mode 有既定的优先级和支持属性范围；没有显式覆盖的属性将按既有解析规则继承。

### 6.5 Preview 预览

预览可用于提前检查 Markdown 结构和支持的公式、Mermaid、链接、表格和样式效果，但 Word 中的分页、页眉页脚、字体度量和复杂字段可能不同。完成转换后仍应检查最终 DOCX/PDF。

## 7. 命令行（CLI）操作与自动化

### 7.1 基本语法

```powershell
& $cli <Markdown文件或目录> --format docx|pdf|both [其他受支持选项]
```

这是便于阅读的语法说明，不能直接逐字运行；请用实际文件路径替换占位符。权威选项列表来自：

```powershell
& $cli --help
```

### 7.2 常见命令

**单个 DOCX：**

```powershell
& $cli ".\report.md" --format docx --no-open
```

**PDF（需要 Word）：**

```powershell
& $cli ".\report.md" --format pdf --no-open
```

**DOCX 和 PDF：**

```powershell
& $cli ".\report.md" --format both --no-open
```

**指定输出目录：**

```powershell
& $cli ".\report.md" --format docx --output ".\exports" --no-open
```

**目录转换：**

```powershell
& $cli ".\markdown-documents" --format docx --no-open
```

**查看帮助：**

```powershell
& $cli --help
```

### 7.3 CLI 常用参数说明

| 参数 | 含义与注意事项 |
|---|---|
| `--format docx` | 生成 DOCX。旧调用未明确指定格式时，默认行为设计上保持 DOCX。 |
| `--format pdf` | 从 DOCX 路线经 Microsoft Word 输出 PDF。 |
| `--format both` | 输出 DOCX 和 PDF，需检查两类结果及部分成功提示。 |
| `--output` / `-o` | 指定输出位置；实际行为以 `--help` 和结果消息为准。 |
| `--no-open` | 不自动打开生成文件；适合测试和脚本。 |
| `--config` / `-c` | 使用受支持的既有配置输入；配置结构应参考项目当前配置文档。 |
| `--verbose` / `-v` | 查看更详细的诊断信息。 |

其他 ASCII/预览相关参数属于高级选项，建议仅按实际 `--help` 输出和受支持的输入格式使用。`--style-config` 并未完成 2.1.0 CLI 验收，**不要在自动化脚本中依赖它**。

### 7.4 自动化示例：保存日志与退出码

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

这是示例脚本，运行前应自行确认源文件、输出目录、磁盘空间、路径权限和日志是否含有敏感内容。当前 CLI 记录中 `0` 代表成功、`1` 代表失败、`3` 代表部分成功；最终以已安装程序实际输出合同为准。

### 7.5 不要在脚本中做的事

- 不要依赖 CLI `--version` 或尚未完成的样式选择参数。
- 不要只凭“文件存在”就判定转换正确。
- 不要把 PDF 请求视为无需 Microsoft Word 的独立渲染器。
- 不要自动重试并覆盖同名的用户文件。
- 不要在无人值守环境中自动关闭用户已有的 WINWORD.EXE 进程。

## 8. DOCX、PDF、Both 的选择和结果判读

### 8.1 何时使用哪种格式？

| 目标 | 推荐选择 | 理由 |
|---|---|---|
| 需要修改正文、公式和样式 | DOCX | 可继续在 Word 中编辑。 |
| 需要向客户提供固定页面的阅读稿 | PDF | Word 原生导出更接近 Word 的显示。 |
| 需要同时保留编辑版和阅读版 | Both | 一次工作流输出 DOCX 与 PDF。 |
| 机器没有 Word | DOCX | 不要选择 PDF 作为必成功的输出。 |

### 8.2 Both 中的部分成功

例如 DOCX 已生成，但 Word 未安装、Word 出错或 PDF 导出失败时，请依据 GUI/CLI 显示的状态和诊断确定保留了哪些 Artifact。不要把 DOCX 的 PASS 外推为 PDF 的 PASS。

### 8.3 DOCX 中的目录与字段

某些 Word 原生功能（例如 TOC 字段、交叉引用字段或分页相关内容）可能需要在 Word 中更新。若看到目录页码或字段尚未刷新，可先保存原文件副本，再在 Word 中更新相关字段并人工复核。

Build5 默认 `word_com=False`，意味着普通 DOCX 转换不会为了 TOC 刷新而启动 Word；这不是 Word COM 可选刷新分支已经完全通过测试的证明。

## 9. 样式预设、Profiles、Templates 与 Project Mode

### 9.1 选择策略

- **Profile：** 快速选定一组既有的文档表现规则。
- **Word Template：** 引入受支持的 `.docx` 模板资源或样式结构。
- **Style Preset：** 只为明确支持的字体、标题、段落等属性提供覆盖。
- **Project Mode：** 将项目层面的选项和多个文档组织成可复用工作流。

它们并非简单的“最后选择项永远覆盖其他所有项”。Finisova 使用现有的 Presentation Resolution 规则解析所支持的属性，用户应通过小样本对目标输出进行确认。

### 9.2 一个稳妥的排版过程

1. 先使用默认样式生成最小 DOCX。
2. 选择合适 Profile，检查标题与正文。
3. 如需企业格式，再测试指定 Word Template。
4. 最后仅覆盖确实需要改变的 Style Preset 属性。
5. 保存或导出可复用预设，并用代表性长文档再次验证。

### 9.3 Project Mode / Batch

对于包含多个 Markdown 文件的项目，推荐先用一小组样本验证：文件选择范围、项目配置识别、图片相对路径、输出位置、相同文件名冲突、失败与部分成功提示。项目清单的精确文件格式及 GUI 菜单入口应以实际产品帮助和接受的 Project Mode 规格为准；本手册不虚构未核实的 Manifest 语法。

## 10. 公式、目录、标题、引用及其他复杂文档

### 10.1 标题、列表与表格

采用标准 Markdown 标题层级：

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

标题和列表尽量保持结构清晰；复杂合并单元格、嵌入 HTML 或任意 CSS 不能保证与 Word 等效。

### 10.2 行内和块级公式

产品实现支持限定范围的 `$...$` 和 `$$...$$` 公式语法，并以 Word 原生可编辑公式为主要 DOCX 表示形式。例如：

````markdown
The equation is $E = mc^2$.

$$
\frac{a+b}{c}
$$
````

不支持的 TeX 命令、宏或扩展语法可能产生提示或回退显示。生成后应在 Word 中检查公式可编辑性与实际版式。

### 10.3 目录（TOC）

对于长篇文档，使用一致的 Markdown Heading 结构。Word 目录以文档结构为基础；目录深度、模板、字段刷新及页码可能影响最终显示。正式交付前在 Word 中检查并在需要时更新目录字段。

### 10.4 Captions 和 Cross References

图、表、公式标题以及交叉引用需要遵循产品实际支持的 Canonical/Reference 语法。不要把一般 Markdown 的图片 Alt Text、任意锚点或普通括号文本直接当作可自动编号的 Word Caption/Cross Reference。推荐从已经通过转换的示例项目复制该语法，并验证交叉引用指向正确的目标。

### 10.5 Citations / Bibliography

受支持的引用语义包含形如 `[@key; @key]` 的键式引用；文末 Bibliography 生成仍取决于项目引用数据、所用 BibTeX 子集及转换配置。不要凭空假定任意 `.bib` 文件都会自动被发现。若有重复/缺失引用键或 BibTeX 无法解析，应先修正源数据再转换。

## 11. 高级预览、Mermaid 与离线资源

### 11.1 预览支持什么？

已接受的 Preview 方向包括公式、Mermaid、链接、表格、标题、代码块等富文本与 Style-aware 展示。预览以 Canonical 语义模型为来源，并对不支持的输入提供受控显示或诊断。

### 11.2 Mermaid 示例

````markdown
```mermaid
flowchart LR
    Draft[Markdown] --> Docx[DOCX]
    Docx --> Pdf[PDF via Word]
```
````

该示例用于演示语法，具体图类型、节点标签和运行时渲染能力仅保证在受支持的 Mermaid 子集内。特别是 Build5 的独立 Windows **安装后离线 Preview 尚未通过完整验收**。

### 11.3 预览与 PDF 的视觉差异

- 预览不是 Word 排版引擎。
- Word 的字体可用性、页边距和分页计算会影响最终结果。
- Word 字段（例如 TOC）可能需要在 Word 中更新。
- SVG 图片嵌入 DOCX 的数据 URI 路径在 Build5 **未完成端到端验证**。

### 11.4 链接与外部网络

预览设计上对链接和资源访问实施约束。点击外部 `https://` 链接可能主动访问互联网；不要把“本地转换”理解为任何操作都永不联网。对安全敏感的资料，不要打开来源不明的外部链接。

## 12. 文件、路径、覆盖保护与批量处理

### 12.1 选择输出位置

优先使用你能读写且容易备份的文件夹。对长路径、受保护目录、只读盘符或当前正在使用的 Word 文件，应保留原始输入并根据错误信息调整输出位置。

### 12.2 相对图片路径

Markdown 图片使用相对路径时，图片解析可能依赖 Markdown 所在位置、项目根目录和实际支持的语法。推荐：

```text
report/
  report.md
  assets/
    chart.png
```

在 Markdown 中：

```markdown
![Chart](assets/chart.png)
```

转换后核对图片是否被正确带入 DOCX；不要将图片缺失当作转换文件本身不存在。

### 12.3 批量转换前检查

- [ ] 先备份原始 Markdown 与引用资源。
- [ ] 使用代表性小样本验证输出。
- [ ] 确认相对图片、BibTeX 和项目配置的路径。
- [ ] 避免同名文件相互覆盖。
- [ ] 确认可用磁盘空间。
- [ ] 阅读逐文件诊断，不要把部分成功当作全面通过。

## 13. 卸载、重装与文件保留

**仅适用于获得明确授权的独立测试电脑。** 当前 Build5 在干净 Windows 上的卸载/重装行为尚未最终验收。

测试时可通过 Windows **Settings → Apps → Installed apps** 寻找 Finisova，使用 Windows 提供的卸载入口。卸载后记录程序目录、快捷方式、卸载条目是否移除，检查用户创建的 Word/PDF 和 Markdown 文件是否仍安全保留。

不要使用脚本清空 `%LOCALAPPDATA%`、`Documents` 或其他目录来“确认卸载”，也不要手工删除不确定归属的用户文件。重装前记录环境信息和安装包 Hash，以便比较。

## 14. 隐私、安全与数据处理边界

Finisova 的目标是让文档在本机处理，不要求将 Markdown 内容上传到外部云转换服务。但某些功能可能涉及系统集成、Microsoft Word、本地渲染组件以及用户主动点击的网络链接。

**不作以下未验证承诺：** 完全无网络流量、无遥测、已通过渗透测试、绝对零残留、全平台兼容或满足特定监管合规标准。正式商业版本需要另行验证。

内部测试时请尽量使用合成样本文档，避免将客户秘密、密码、个人识别信息、商业机密或未经授权的文件附在问题报告里。

## 15. 常见问题与故障排查

| 现象 | 可能原因 | 处理建议 |
|---|---|---|
| PowerShell 提示找不到 `finisova-cli.exe` | 安装目录未在 PATH | 使用本手册中的完整 CLI 路径。 |
| PDF 报告缺少 Microsoft Word | Word 不可用或未正确配置 | 先改用 DOCX；在获授权测试机核验 Word。 |
| DOCX 生成成功但 PDF 失败 | Both 的 PDF 分支失败 | 读取部分成功/诊断信息，勿误报全部 PASS。 |
| 输出文件找不到 | 输出目录与自动命名不同于预期 | 查阅转换结果提供的实际路径。 |
| 文件无法写入或被锁定 | 文件夹无权限或文件正被 Word 打开 | 换可写目录；避免覆盖正在打开的文件。 |
| 公式或 Mermaid 预览与 Word 不同 | 两种渲染路径的能力和版式不同 | 将生成的 DOCX/PDF 作为最终检查对象。 |
| TOC 页码不正确 | Word 字段未刷新 | 备份后在 Word 中更新相关字段。 |
| SVG 或图表未出现 | 特定路径/语法未支持或未验证 | 记录最小复现 Markdown，不推断已支持全部 SVG。 |
| 安装被 Defender/SmartScreen 警告 | 内部未签名候选、信誉或安全策略 | 不关闭安全防护；保存错误供 QA 审查。 |
| CLI `--version` 不被识别 | 此选项不是接受的 CLI 合同 | 改用 `--help`，通过安装包身份确认版本。 |
| 预览或退出时发生崩溃 | GUI/Worker 环境或原生组件问题 | 保留退出码、系统版本、复现步骤并报告。 |

### 15.1 故障报告应包含什么？

1. Windows 版本、Microsoft Word 是否安装及版本信息（如可获取）。
2. Installer SHA-256、安装位置和是否处于离线环境。
3. GUI/CLI/Preview/安装/卸载的具体步骤。
4. 输入 Markdown 的可公开**脱敏最小样例**。
5. 所选输出格式、Profile、Template、Style Preset。
6. 实际错误、Exit Code、诊断文本及生成文件是否有效。

## 16. 验证状态、已知限制与版本身份

### 16.1 已记录的 Build5 工程验证

| 证据范围 | 结果 |
|---|---|
| GUI / CLI PyInstaller Build | PASS |
| Inno Setup Installer Build | PASS |
| Qt 组件 D-1、License Metadata D-2、Native Provenance D-3 | PASS |
| Controlled Cairo Runtime | 19/19 DLL，Hash 一致 |
| GUI 开关窗口、CLI 帮助、CLI DOCX Smoke | PASS（开发机） |
| Packaging Focused Tests | 36/36 PASS |
| Bounded Integration Tests | 42/42 PASS |
| 已执行验证范围内的新增失败 | 0 |
| Post-Build Audit | ACCEPTED WITH DISCLOSED QUALIFICATIONS |

### 16.2 尚未完成的 QA 或合规项目

| 项目 | 状态 |
|---|---|
| 另一台 Windows 的 Clean Install / Uninstall / Reinstall | NOT VERIFIED |
| 不安装 Word 的独立电脑 DOCX / PDF 错误行为 | NOT VERIFIED |
| 在已安装环境中的离线 Math / Mermaid Preview | NOT VERIFIED |
| SVG 数据 URI → DOCX 嵌入图片 | NOT VERIFIED |
| 可选 Word COM 字段刷新及完整所有权清理路径 | 尚非默认 Smoke 所证实的范围 |
| 商业第三方许可证 §19A–§19E | OPEN |
| P24-03 Parent Closure / P24-04 RC Freeze / Production Release | NOT AUTHORIZED |

**法律和发行说明：** 随包 EULA 与第三方许可证文件不代表商业再分发审查已经完成。试用期、许可证激活、价格、升级权限及公开下载页面也未在本手册中被确认。

### 16.3 内部版本核验

```text
Product        : Finisova 2.1.0
Candidate      : Build5 (internal only)
Installer name : Finisova_v2.1.0_Setup.exe
Installer size : 114101262 bytes
SHA-256        : 2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d
Build baseline : c515974614d6323abadd5e14217b73a7387b3416 + hashed overlay revisions
```

**可重现性边界：** Build5 是 Git HEAD 加经哈希记录的商业源码 Overlay 和受控 Native Runtime 形成的构建；**不能仅凭该 HEAD 在一个全新仓库中得到相同二进制文件**。

## 17. 术语表、反馈模板与测试核对表

### 17.1 常用术语

| 术语 | 含义 |
|---|---|
| DOCX | 可以在 Microsoft Word 中编辑的文档格式。 |
| PDF | 更适合固定版式分发与阅读的文档格式。 |
| Both | 在同一请求中输出 DOCX 与 PDF。 |
| Profile | 文档外观/配置的预定义选择。 |
| Style Preset | 受支持属性的可复用样式预设。 |
| Word Template | Word 模板文档或样式资源。 |
| TOC | Table of Contents，目录。 |
| Caption | 图、表或公式的标题与编号结构。 |
| Cross Reference | 对其他标题/对象的引用。 |
| Citation/Bibliography | 文内参考引用与文末文献列表。 |
| Artifact | 某一输出请求生成的 DOCX、PDF 等实际文件。 |
| Internal Candidate | 仅供授权内部验证的候选构建，不等于公开发布版。 |

### 17.2 最小故障报告模板

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

### 17.3 内部安装测试基本核对表

- [ ] 测试目标机器已获明确授权，且无用户生产数据风险。
- [ ] Installer 原始 Hash 已核验。
- [ ] 安装位置、快捷方式和发行者信息已记录。
- [ ] GUI 与 CLI 从**安装目录**运行，不是源代码目录。
- [ ] DOCX 转换和文件可读性已验证。
- [ ] Word-present PDF/Both 场景已按需要单独验证。
- [ ] Word-absent 场景在独立环境真实执行；未通过卸载用户的 Word 来制造。
- [ ] 安装后的 Offline Preview 已在需要的场景检查。
- [ ] 卸载和重装没有删除用户资料。
- [ ] 已注明未完成的 SVG、§19 许可和其他已知限制。

---

**Finisova — 将 Markdown 整理为专业的可编辑 Word 文档。**  
**文档说明：** 本手册基于截至 2026-10-16 的工程规格、已接受的 Build5 记录和 Post-Build Audit 编写。受支持的按钮标签、安装兼容性及正式授权条款应以未来获准的正式发行包与经核对的产品帮助为准。
