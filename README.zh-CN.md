# Finisova 2.1.0 — 中文说明

**将 Markdown 转换为专业 Word 文档：本地处理、重视隐私、不依赖订阅式云转换。**

[English README](README.md)

> **重要状态声明：** Finisova 2.1.0 Build5 目前仅是**已冻结的内部测试候选版**，不是正式公开发布的软件。虽然已通过开发电脑上的 GUI/CLI、构建及有限集成验证，但独立 Windows 安装、Word 缺失环境、离线预览及第三方许可合规审核仍未完成。未经授权，不应把安装包上传到公开 GitHub Releases。

- 产品：**Finisova**
- 版本：**2.1.0 Build5**
- 平台：**Windows 桌面**
- 安装包内的 Publisher 标识：**Arculoric**
- 当前状态：**Internal Evaluation Only**

## 1. 产品用途

Finisova 让用户保留 Markdown 的快速写作体验，同时生成可编辑、适合交付的 Microsoft Word 文档。支持的典型场景包括技术报告、研究文档、项目记录、AI 生成的 Markdown 草稿和长篇文档。

核心转换路线：

```text
Markdown → Canonical 解析与排版解析 → DOCX
                                          └→ Microsoft Word 原生 PDF 导出
```

PDF 是从 DOCX 经 Microsoft Word 导出，并非独立的 Markdown-to-PDF 渲染器。

## 2. 功能概览

| 功能 | 2.1.0 范围与说明 |
|---|---|
| Markdown → DOCX | 不依赖 Microsoft Word 生成 DOCX |
| PDF / Both 输出 | 需要安装并正确配置 Microsoft Word |
| 图形界面 | 文件选择、输出格式、状态与诊断 |
| CLI | 单文件、目录及 `--format docx/pdf/both` |
| 专业排版 | Profiles、Word Templates、GUI 样式预设 |
| Word 文档结构 | 标题、表格、目录、元数据及受支持的文档元素 |
| 技术及学术文档 | 可编辑公式、图表标题、交叉引用、文内引用与参考文献 |
| Project Mode / Batch | 项目与批量处理场景 |
| 高级预览 | 公式、Mermaid、链接、富内容及样式感知预览 |

**已知范围限制：** CLI 样式预设参数尚未完成；预览不是 Word/PDF 的像素级所见即所得；Build5 中 SVG→DOCX 端到端嵌入路径未验证；LibreOffice 不是 2.1.0 PDF 后端。

## 3. 安装前的要求

当前仅面向已经获得授权的内部测试电脑。独立 Windows 11 环境是下一阶段的验证目标，尚不能把它表述为已经通过完整兼容性验收。

已冻结 Build5 文件：

`Finisova_v2.1.0_Setup.exe`

SHA-256：

```text
2674704f7f3627e7c05c10728e223b5a83e4ca6a3e2bdbc2dc1f70cf0e2c223d
```

PowerShell 验证：

```powershell
Get-FileHash ".\Finisova_v2.1.0_Setup.exe" -Algorithm SHA256
```

当前 Installer 的设计为当前用户安装，预期目录：

```text
%LOCALAPPDATA%\Programs\Finisova
```

Installer 不会自动修改系统 PATH；正式安装与卸载仍须在另外授权的测试环境中核验。

## 4. GUI 使用方法

1. 启动 **Finisova**。
2. 选择 Markdown 文件，或进入受支持的输入/项目工作流。
3. 选择 `DOCX`、`PDF` 或 `Both`。
4. 选择输出位置及需要的 Profile、Template 或样式预设。
5. 可先查看预览，但必须理解它不保证 Word 的分页与字体完全一致。
6. 启动转换，核对生成文件和所有诊断消息。

第一次测试建议使用 `DOCX`。PDF/Both 需在拥有 Microsoft Word 的测试环境中执行。

## 5. CLI 使用

安装后预期 CLI 路径：

```powershell
$cli = Join-Path $env:LOCALAPPDATA "Programs\Finisova\cli\finisova-cli.exe"
& $cli --help
```

转换为 DOCX：

```powershell
& $cli ".\example.md" --format docx --no-open
```

转换为 PDF（需要 Word）：

```powershell
& $cli ".\example.md" --format pdf --no-open
```

同时生成 DOCX 和 PDF：

```powershell
& $cli ".\example.md" --format both --no-open
```

指定输出目录：

```powershell
& $cli ".\example.md" --format docx --output ".\exports" --no-open
```

目录处理：

```powershell
& $cli ".\markdown-documents" --format docx --no-open
```

注意：`--version` 不是当前已接受的 CLI 参数。GUI 支持样式预设管理，但 CLI 样式预设选择尚未完成。以程序内 `--help` 为参数权威来源。

## 6. Word 与 PDF 的关系

- **DOCX 模式：** 设计上无需 Microsoft Word。
- **PDF 模式：** 需要可用的 Microsoft Word，通过其原生导出能力生成 PDF。
- **Both 模式：** 使用 DOCX-first 路线生成 DOCX 与 PDF。
- **无 Word 环境：** PDF 请求应给出明确依赖诊断，但 Build5 在独立无 Word 测试机上的实测仍待完成。
- **LibreOffice：** 2.1.0 已明确延期，不支持将其作为 PDF 替代后端。

## 7. 隐私和离线处理

Finisova 是本地桌面文档转换软件，而非要求上传源文件的云服务。通常的 Markdown→DOCX 流程在本机执行。点击外部 HTTP/HTTPS 链接属于另一类可能访问网络的操作。

这不是对 Build5 已完成完整网络隔离、无遥测或信息安全审计的声明。正式发布前仍须核验安装包的离线运行及依赖行为。

## 8. 已知限制与正式发行前任务

| 项目 | 状态 |
|---|---|
| Build5 本地 GUI/CLI/Installer 构建及 Smoke | PASS（记录范围内） |
| 36 项 Packaging 检查、42 项 Integration 检查 | PASS |
| 独立 Windows 干净安装 | NOT VERIFIED |
| 安装后的 Word-absent 场景 | NOT VERIFIED |
| 安装后的离线公式/Mermaid 预览 | NOT VERIFIED |
| SVG 数据 URI 转为 DOCX 嵌入图片 | NOT VERIFIED |
| CLI 样式预设选择 | 未完成 |
| Third-party Licensing §19A–§19E | OPEN |
| RC Freeze / Production Release | NOT AUTHORIZED |

Build5 默认采用 `word_com=False`；普通 CLI DOCX Smoke 未创建 Word 进程。但不能据此认定可选 Word COM 刷新分支的完整进程清理行为已通过验证。

## 9. 常见问题

**Q：不安装 Microsoft Word，能生成 DOCX 吗？**  
设计上可以。Build5 已在开发机通过 DOCX CLI Smoke；独立无 Word 测试机仍需要验收。

**Q：为什么不能导出 PDF？**  
确认已安装可用的 Microsoft Word。LibreOffice 不能替代本版本的 Word 原生 PDF 后端。

**Q：为何 PowerShell 找不到 `finisova-cli.exe`？**  
Installer 不修改 PATH，请使用完整安装路径调用。

**Q：为什么预览与 Word 文件排版不同？**  
预览是语义级预览，不保证 Word 级分页、行距、字体度量、复杂图片与表格渲染完全一致。

**Q：可以直接公开上传 Build5 安装包吗？**  
不可以。Build5 尚未取得第三方许可合规、安装验收和 Human Release Authority。

## 10. 许可与发布

产品使用适用相应 `EULA.txt`；第三方组件信息见 `THIRD_PARTY_NOTICES.txt` 和随包许可证材料。**公开发行所需的 §19 合规审核尚未完成。**

不要根据本 README 推定已经确定 14/30 天试用、License Key、正式定价、升级权益或订阅/永久许可的法律条款。

源码仓库：[Quansheng2022/finisova](https://github.com/Quansheng2022/finisova)。仓库中出现代码或安装包，不代表产品已获准公开发行或源码具有特定开源许可。

---

**Finisova 2.1.0 Build5 — 内部工程候选已冻结；Clean Windows QA、合规审核及生产发布仍属于独立审批门槛。**
