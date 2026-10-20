# Zotero Word AutoCite

[English](README.md) | **简体中文**

用 Zotero 的 Word 加载项插入一条引用很快；按引用方案在整篇稿件中放好几十条引用，就费时多了。这个 Codex Skill 负责这一步：将选定文献对应到指定位置，通过 Zotero 的 Word 集成插入或修改字段，并刷新、核验参考文献表。写入的引用仍可在加载项中编辑和刷新。

> [!NOTE]
> **Human-in-the-Loop.** 这个 Skill 承担重复性的科研事务，是辅助研究，而不是替代研究。问题、判断和结论留给人，每个结果使用前都要检查。

<img src="assets/overview.png" alt="引用方案、文献（Zotero 条目或 PDF）和 Word 稿件交给 AI Agent，依次完成文献匹配、插入 Zotero 字段、在 Word 中刷新、核验参考文献表；结果是带可刷新 Zotero 引用字段和参考文献表的 Word 文档，已有引用可以继续调整并再次刷新。" width="1600">

引用方案可以是笔记、表格（Markdown、CSV、JSON、Excel）或 Word 批注；文献可以是 Zotero 条目，也可以是附元数据的 PDF。哪些文献支撑哪些论述、顺序如何，由作者决定；方案不完整时，工作流报告缺口而不代为补全。文献检索与下载见 [ScholarFetch](https://github.com/LiX-Works/ScholarFetch)。

## 任务

| 任务 | 输入 | 结果 |
|---|---|---|
| 只读检查 | 问题或文件 | 检查结论 |
| 整理引用输入 | 已有决定、笔记、表格或 Word 批注 | 引用任务表与文献映射 |
| 准备 Zotero | 已核实的记录、PDF 或已有分类 | 按授权形成导入计划或准备条目 |
| 插入引用 | Word、引用方案、文献来源与样式 | 含 Zotero 字段和参考文献表的工作副本 |
| 调整引用 | 已有 Zotero 字段的 Word 与修改方案 | 定点修改，其余引用保留 |
| 刷新与核验 | 已有 Zotero 字段的 Word | 沿用现有样式的刷新副本 |

## 开始使用

需要 Windows、Word 桌面版、Zotero 及其 Word 加载项、Python 3 和 PowerShell；新文档还需指定引用样式。写入前，工作流会检查 Zotero 本地 API 和 Word 集成 DLL，但不负责安装。

仓库根目录即 Skill 目录。安装为 Codex 个人 Skill：

```powershell
git clone https://github.com/LiX-Works/Zotero-Word-AutoCite.git (Join-Path $env:USERPROFILE '.codex\skills\zotero-word-autocite')
```

请求示例：

```text
使用 $zotero-word-autocite。项目在 C:\path\to\project，
目标文件是 manuscript.docx。

请根据项目中的引用方案和文献来源，在指定位置插入 Zotero 引用字段。
使用我已选定的引用样式，完成后刷新并检查参考文献表。
```

修改已有文档时，说明涉及的段落和文献变更；仅刷新不需要引用方案。Skill 与工作流文档为英文。

## 安全机制

- 修改在工作副本上进行，源文档保留。
- 写入前核对文献身份。文献身份或位置不明确、输入冲突或保存状态不确定时，相应操作停止。
- 批量脚本（`build_payload.py`、`insert_zotero_fields.ps1`）用于新增引用；替换或删除已有引用组采用定点字段修改。
- Zotero 导入需经授权并先备份数据库。默认不删除条目、PDF 或已有引用。
- 只有在保存后的文档中核对字段、引用条目和参考文献表之后，任务才算完成。

## 文档与测试

| 文件或目录 | 用途 |
|---|---|
| [SKILL.md](SKILL.md) | 任务指令与入口 |
| [工作流总览](workflow/README_FIRST.md) | 选择检查、准备、插入、调整或刷新路线 |
| [输入整理](workflow/PREPARE_CITATION_INPUTS.md) | 理解已有材料，形成任务表和文献映射 |
| [Word–Zotero 工作流](workflow/ZOTERO_WORD_AUTOMATION_WORKFLOW.md) | 字段操作、刷新、排错与验证 |
| [脚本](workflow/zotero_word_automation/scripts/)与[模板](workflow/zotero_word_automation/templates/) | 辅助操作与输入格式示例 |
| [测试](workflow/zotero_word_automation/tests/) | 离线回归测试 |

```powershell
python -m unittest discover -s workflow/zotero_word_automation/tests -v
```

离线测试不启动 Word 或 Zotero，不能证明特定电脑上的实际插入和刷新可用。

采用 [MIT 许可证](LICENSE)。
