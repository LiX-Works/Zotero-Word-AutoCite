# Zotero Word AutoCite

[English](README.md) | **简体中文**

一份按你的引用方案，在 Microsoft Word 中写入可刷新 Zotero 字段的 Codex Skill。

写论文时，文献可能散在 PDF 文件夹、Zotero 文献库和几份表格里，引用位置又分布在 Word 的不同段落。这个工作流把已有材料对齐，将你选定的文献对应到指定位置，插入或调整 Zotero 引用字段，再刷新并检查参考文献表。

交付的引用仍可通过 Zotero 的 Word 加载项继续编辑；写入的是可刷新的字段，而非手工输入的数字编号。

首页提供中英双语；Skill 与详细工作流文档使用英文。

## 选择你需要的任务

工作流有多个入口，不必每次都跑完整流程。

| 任务 | 需要什么 | 交付什么 |
|---|---|---|
| 只读检查 | 待检查的问题或文件 | 检查结果或解释 |
| 整理引用输入 | 已有决定、笔记、表格或 Word 批注 | 引用任务表、文献映射和交接说明 |
| 准备 Zotero | 已核实的文献记录、PDF 或已有分类 | 按授权形成导入计划，或准备条目并核验映射 |
| 为新文档插入引用 | Word、引用方案、文献来源和引用样式 | 含 Zotero 字段、按方案需要生成的参考文献表和验证记录的工作副本 |
| 调整已有引用 | 已有 Zotero 字段的 Word 与具体修改方案 | 局部变更，保留未受影响的引用 |
| 仅刷新与验证 | 已有 Zotero 字段的 Word | 沿用现有样式刷新，并交付核验结果 |

引用方案可以来自 Markdown、CSV、JSON、Excel、Word 批注或文字说明。文献可以已在 Zotero 中，也可以先放在 PDF 文件夹并附有元数据。Codex 会先核对已有信息，不要求你提前统一成模板。

哪些论文支撑哪些论述、顺序和落点，由你确定。缺少这些信息时，工作流可以整理材料和指出缺口，但不会代你决定最终引用哪些文献。

## 需要准备什么

- Windows、Microsoft Word 桌面版、Zotero 和 Zotero Word 加载项。
- Python 3 与 PowerShell。
- 目标 Word、文献来源，以及希望执行的引用方案或修改内容。
- 新文档需要指定引用样式；仅刷新已有文档时可以沿用当前样式。

实际操作字段前，工作流会核对本机环境，包括 Zotero 本地 API 和官方 Word 集成 DLL。Skill 不代替这些应用的安装与配置。

## 安装与第一次使用

本仓库根目录就是 Skill 目录：`SKILL.md` 与 `workflow/` 位于同一层。安装时保留完整结构。

可以在 PowerShell 中克隆到 Codex 的个人 Skill 目录：

```powershell
git clone https://github.com/LiX-Works/zotero-word-autocite.git (Join-Path $env:USERPROFILE '.codex\skills\zotero-word-autocite')
```

也可以下载仓库，把完整文件夹放到该位置。已有同名目录时，先比较个人改动；如果 Codex 没有显示新 Skill，可重新打开任务后检查。

为新文档插入引用时，可以这样开始：

```text
使用 $zotero-word-autocite。项目在 C:\path\to\project，
目标文件是 manuscript.docx。

请根据项目中的引用方案和文献来源，在指定位置插入 Zotero 引用字段。
使用我已选定的引用样式，完成后刷新并检查参考文献表。
```

调整已有文档时，说明受影响的段落和准确的文献变更。例如，用周围原句定位某一引用组，按已确认的文献表指定新成员，并要求保留文档其余引用。仅刷新时，无须重新建立完整引用方案或映射表。

## 修改与结果如何处理

编辑任务默认保留源文档，在项目内的工作副本上操作。明确指定的可修改副本可以复用，再次修改前先备份。写入前核对输入和文献身份。

你已经要求完整执行且检查通过时，Codex 会在已有授权内继续相关阶段；文献身份或落点无法确定、信息严重冲突、文档损坏或保存状态不明时，暂停受影响的操作并询问。只读检查或方案任务到相应交付为止。

批量脚本 `build_payload.py` 和 `insert_zotero_fields.ps1` 负责**新增引用**。替换或删除已有引用组，按工作流定位并修改相应字段，不用再次追加来代替替换；其他引用组保留。

PDF 清单和导入计划脚本负责准备文件与映射，本身不执行 Zotero 导入。实际导入或条目修改按任务授权进行，并备份文献库数据库；不默认删除条目、PDF 或清空全部 Word 引用。

完成时检查实际字段、引用成员、参考文献表和保存后的文档。脚本成功退出或等待结束，不能单独证明 Zotero 刷新完成。交付包括工作副本，以及已验证内容和未完成事项的简短说明。

## 详细文档与离线测试

| 文件或目录 | 用途 |
|---|---|
| [SKILL.md](SKILL.md) | 任务指令与入口 |
| [工作流总览](workflow/README_FIRST.md) | 选择检查、准备、插入、调整或刷新路线 |
| [输入整理](workflow/PREPARE_CITATION_INPUTS.md) | 理解已有材料，形成任务表和文献映射 |
| [Word–Zotero 工作流](workflow/ZOTERO_WORD_AUTOMATION_WORKFLOW.md) | 字段操作、刷新、排错与验证 |
| [脚本](workflow/zotero_word_automation/scripts/)与[模板](workflow/zotero_word_automation/templates/) | 辅助操作与输入格式示例 |
| [测试](workflow/zotero_word_automation/tests/) | 离线回归测试 |

在仓库根目录运行：

```powershell
python -m unittest discover -s workflow/zotero_word_automation/tests -v
```

这些测试不会启动 Word 或 Zotero，也不能证明某台电脑上的真实插入和刷新已经通过。

本项目采用 [MIT 许可证](LICENSE)。
