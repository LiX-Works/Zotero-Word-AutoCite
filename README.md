# Zotero Word AutoCite

做论文时，文献可能散在 PDF 文件夹、Zotero 和几份表格里，引用位置又分布在 Word 的不同段落。这个 Codex Skill 会按你已经确定的引用方案，把文献和位置逐一对好，插入真正的 Zotero 引用字段，然后刷新并检查参考文献表。以后仍可以在 Word 里用 Zotero 继续编辑这些引用。

它可以接手几类工作：从零给文档插引用；调整已有引用；只刷新并检查一份文档；或者先把格式不统一的文献表和引用计划整理好。遇到论文身份、引用位置等关键问题无法确定时，它会停下来问你。

## 需要准备什么

- Windows、Microsoft Word 桌面版、Zotero 和 Zotero 的 Word 加载项。
- Python 3、PowerShell。
- 目标 Word，以及你希望哪些位置引用哪些文献的方案。文献可以在 Zotero 中，也可以先放在 PDF 文件夹并附有元数据；方案可以是表格、Markdown、批注或文字说明。

新文档还需要确定引用样式；已有文档如果只是刷新，可以沿用当前样式。

## 安装与使用

在 PowerShell 中把仓库克隆到 Codex 的个人 Skill 目录：

```powershell
git clone https://github.com/LiX-Works/zotero-word-autocite.git (Join-Path $env:USERPROFILE '.codex\skills\zotero-word-autocite')
```

在 Codex 中可以这样开始：

> 使用 `$zotero-word-autocite`。项目在 `C:\path\to\project`，目标是 `manuscript.docx`。请根据项目里的引用方案和文献文件夹插入 Zotero 引用，完成后刷新并检查参考文献表。引用样式为我指定的样式。

已有文件和对话里能找到的信息，Codex 会先自行核对；只有关键内容缺失或互相冲突时才会集中询问。

## 工作流与脚本

[Skill 入口](SKILL.md)会按任务指向[输入整理说明](workflow/PREPARE_CITATION_INPUTS.md)和[Word–Zotero 工作流](workflow/ZOTERO_WORD_AUTOMATION_WORKFLOW.md)。脚本、模板和离线测试在 `workflow/zotero_word_automation/`。

批量插入脚本负责新增引用。调整已有引用时，Codex 按工作流定位并修改对应的 Zotero 字段，再统一刷新和核验。PDF 清单和 Zotero 导入计划也有脚本辅助；实际导入会根据当前文献库处理。

离线测试可在仓库根目录运行：

```powershell
python -m unittest discover -s workflow/zotero_word_automation/tests -v
```

本项目采用 [MIT License](LICENSE)。
