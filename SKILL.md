---
name: zotero-word-autocite
description: 在 Windows 的 Microsoft Word 与 Zotero 中整理引用输入、准备文献映射和导入、插入或调整可刷新的 Zotero 引用字段、刷新并验证参考文献表。用于 Zotero-Word 引用自动化、已有引用增删改、引用任务表适配或相关故障检查；文献检索和学术选文另行处理。
---

# Zotero Word AutoCite

先读取 [工作流入口](workflow/README_FIRST.md)，按本次任务选择路线，只加载相关说明。

## 输入与执行

从用户对话和项目已有文件确认目标 Word、引用方案、文献来源、样式和操作范围。方案可来自 Markdown、CSV、批注或文字说明，不要求用户预先整理成标准表格。

- 需要适配输入时，阅读 [输入整理说明](workflow/PREPARE_CITATION_INPUTS.md)，形成或复用任务表、crosswalk 和交接记录。
- 需要准备 Zotero、写入或调整字段、刷新验证时，阅读 [技术工作流](workflow/ZOTERO_WORD_AUTOMATION_WORKFLOW.md) 的对应章节。
- 只读检查按问题读取并回复；仅刷新已有文档可沿用已有样式，不要求补齐整套引用任务表。
- 用户要求完整执行且检查通过时，连续完成适用阶段，无独立 dry-run 确认关卡；沿用已有授权。

任意阶段出现严重信息缺失、影响正确性的冲突、文献身份或引用落点无法确定、文档损坏、保存状态不明时，暂停受影响操作并集中询问。普通且可可靠解决的问题自行处理。

保留人工确定的选文、顺序和支撑范围。遵循工作流中的原材料保护、备份与身份核验规则，不以当前 Word 数字编号代替文献身份，不默认插入 backup、not_use 或 manual_gate。

## 工具与结果

脚本位于 [zotero_word_automation/scripts](workflow/zotero_word_automation/scripts/)，模板位于 [templates](workflow/zotero_word_automation/templates/)。以本 Skill 目录为起点解析这些路径；任务产物保存在用户项目中。

执行环境为 Windows、桌面 Microsoft Word、Zotero 及其 Word 加载项。实际字段操作前按技术工作流核验本地 API 和 DLL。

`build_payload.py` 与 `insert_zotero_fields.ps1` 负责新增；按组替换或删除采用技术工作流规定的项目级步骤，不把 replace/delete 当作 add 直接追加。Zotero 导入脚本负责生成计划，实际导入按现有授权和可用接口执行。

完成后核验字段、实际引用成员、参考文献表及保存结果。脚本成功退出或等待结束不能单独证明 ZoteroRefresh 完成。按任务交付整理后的输入，或更新的 Word 工作副本及简短验证结果；说明未完成事项。

`workflow/` 随本 Skill 一起分发；脚本和模板的相对路径在不同项目中保持不变。遇到 Zotero 或 Word 版本差异，先检查实际环境再执行。
