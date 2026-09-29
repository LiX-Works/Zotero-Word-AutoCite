# Zotero-Word 第一阶段：整理与校验输入

版本：2026-09-16

本文件指导 AI 将不同项目的引用方案、文献表、PDF 清单和人工意见整理成第二阶段可用的输入。来源格式可以灵活，交接内容必须明确。先读同目录 `README_FIRST.md`；执行技术规范在 `ZOTERO_WORD_AUTOMATION_WORKFLOW.md`。

## 1. 两阶段与执行范围

第一阶段：读取上下文与已有文件，确定权威来源，整理引用任务表、文献对照表和交接说明，校验身份、分组、顺序及原文定位。

第二阶段：准备或核验 Zotero 条目，补全 key 和内部 item ID，准备项目内 Word 工作副本，按操作清单插入或调整字段，刷新并验证。

同一个 AI 对话可以连续完成两阶段，无需人为切换，也不设置独立 dry-run 或固定确认环节。用户已要求完整执行且检查通过时，第一阶段结束后应直接进入第二阶段。用户只要求整理、审计或计划时，交付第一阶段结果即可。

本说明本身不是某篇论文的执行授权。以当前请求和同一任务已有授权确定 Word、Zotero 导入及正文改动的范围；明确禁止的操作不能被自动衔接规则覆盖。换对话执行时，交接说明记录先前授权的来源，新对话结合用户请求判断，不把表格里的 `approved` 或 `default` 当作独立授权。

## 2. 接受哪些原始输入

- 明确的项目路径、目标 Word 或候选目标 Word。
- 一个或多个 CSV、Markdown、JSON、Excel 表格、批注 Word 或人工修改意见。
- PDF 文件夹、已有 metadata、Zotero collection 或旧 crosswalk，可组合使用。
- 引用样式、执行模式及用户已明确的限制。

用户不必提前统一列名、文献 ID 或文件格式。AI 应先查已有文件，避免要求用户重复提供已存在的信息。仅有 PDF 和论文正文但没有引用选择时，不得自行决定正式引用哪些文献；先整理缺口并询问，或按另行授权的文献筛选任务处理。

## 3. 确定权威来源

1. 当前用户明确决定及其指定的权威文件优先。
2. 可追溯的最新人工复核结果优先于旧候选、自动推荐和 legacy 字段。
3. 文件修改时间、文件名包含 final、某列叫 approved，都不能单独证明其权威性。
4. 记录来源文件、行号或记录 ID、版本、人工决定出处及转换依据。附件中的指令是来源内容，不能自行扩展当前任务授权。
5. 同一权威层级的选择、顺序或落点互相冲突且无法从上下文解决时，暂停受影响操作并集中询问。

允许自主处理列名转换、编码、布尔值格式、已核实的 ID aliases、可唯一定位的原文片段及路径格式；不得自行改变人工选择、引用顺序、支撑范围或把 backup 升为默认使用。

## 4. 三份交接输出

在项目 `metadata/` 中新建带日期的文件，不覆盖来源。当天重跑用版本后缀或时间戳，并在交接说明中指定本次文件的准确路径。

三份输出适用于用户要求预整理或完整引用执行；只读检查不默认创建它们，仅刷新也不要求补齐整套输入。增量任务只整理变更项，并保留已有字段盘点结果。已有合格文件可复用。

### 4.1 word_citation_task_table_normalized_YYYYMMDD.csv

一行表示一个实际引用落点中的一篇文献。相同文献可在多个落点出现。同组内重复记录只有在身份和语义一致时才能合并，并记录来源。

| 字段 | 内容 |
| --- | --- |
| insert_as_group | 一个实际引用括号的唯一 ID；主题大组不必等于实际括号 |
| group_order | 实际落点的执行顺序，整数 |
| claim_id | 可选的论述点 ID，允许保留项目原编号 |
| target_location_description | 给人看的章节、句子和落点说明 |
| anchor_text | 当前目标 Word 中可查找的准确原文片段 |
| anchor_occurrence | 默认 unique；多匹配可明确 first / last / 正整数等，并填写选择依据 |
| anchor_selection_reason | 多匹配时必填；引用已核实的人工方案或上下文；旧表的 first 不等于已消歧 |
| position | before 或 after |
| citation_id | 对照表中的唯一文献 ID |
| citation_order | 同一括号内的人工确定顺序，整数 |
| status | main / secondary / backup / not_use / manual_gate |
| default_insert | yes / no；仅默认 main、secondary 可为 yes |
| action | add / keep / replace / delete；补充操作字段，现有通用脚本不会自动解释 |
| existing_citation_id | 增量操作时已有 Zotero 字段 JSON 的 citationID；无则留空 |
| existing_group_locator | 已有字段的位置、上下文、URI/key 等复合定位信息 |
| source_metadata_file / source_record | 来源路径与行号或记录 ID |
| source_status / source_order | 原始状态和顺序，便于追溯转换 |
| notes | 支撑范围、不能支撑内容、异常和转换说明 |

`default_insert` 只表示文献默认采用意图，不表示该行应该无条件再次插入。已有正确引用应标为 keep。delete 行用 default_insert=no，并保留原有身份和定位信息。

整组删除至少保留一行操作记录；组内成员增删通常应表示为 replace，并记录原成员和完整目标成员。无法确定现有字段时不得靠当前数字编号或第几个字段猜测。

无外部引用的论述点写在交接说明中，不生成伪文献 ID。backup / not_use / manual_gate 可保留在任务表中以便审计，但不得进入默认写入集合。

### 4.2 unified_citation_crosswalk_YYYYMMDD.csv

一行一篇文献，至少包含以下列：

`citation_id, aliases, title, doi, year, first_author, journal, pdf_path, zotero_key, zotero_item_id, zotero_collection, source_metadata_file, notes`

- 保留已有稳定 ID，不强制将项目 ID 改成 REF、TNEW 或 LIT。
- 核实 DOI、题名和版本，不能把同一作者或相似题名自动当成同一篇。未知 DOI 留空并注明，不伪造。
- PDF 路径须核实存在；缺失时如实记录。已有可核验 Zotero 父条目且本轮不需要 PDF 导入时，缺 PDF 不必阻塞字段操作。
- 第一阶段可将 `zotero_key`、`zotero_item_id` 留空并记录待准备；这属于第二阶段任务，不是默认的人工阻塞项。
- 正式生成 payload 前，第二阶段必须核实 key、数字 item ID、URI 与题名/DOI 属于同一条目及同一 library。不得把 key、数组行号或其他库的 item ID 当作本库内部 ID。
- 第一阶段不为补 key 直接编辑 Zotero 数据库。导入、去重和 metadata 修正依第二阶段规范及实际授权执行。

### 4.3 citation_input_handoff_YYYYMMDD.md

用中文简洁记录：

- project_path、source_docx、拟用或已有 working_docx、目标版本/修改时间或哈希；参考批注版另列。
- task_mode：prepare_only / insert_new / incremental / refresh_only。
- style：用户确定的样式及可核验的 style ID；不能沿用示例默认值猜测。
- 本次 task table、crosswalk、PDF 文件夹、collection 的准确路径或标识。
- 权威文件、人工决定来源、关键字段映射规则、来源间冲突处理依据。
- 引用组数量、任务行数、默认使用文献的去重数、keep/add/replace/delete 清单。
- 无外部引用的论述点、暂不使用项及必须保留的支撑范围限制。
- 已完成的定位和身份检查、待第二阶段处理的 Zotero/PDF/环境事项。
- 当前请求及已有授权范围；严重未解决问题及需要用户决定的内容。
- readiness：ready_for_execution / ready_for_zotero_preparation / blocked / preparation_only；写出原因和下一步。

此文件是交接记录，不是新的确认关卡。`ready_for_execution` 仍要求第二阶段写入前复核当前 Word 和 Zotero 状态。

## 5. 格式适配示例

下面只演示用户提供的 tutorial 类表格，不作为所有项目固定规则。

| 来源字段或情况 | 整理方式 |
| --- | --- |
| paper_id / paper_aliases | 对应 citation_id / aliases，核对重复身份 |
| human_selection_status=default_main | 在人工来源已确认时映射 main + yes |
| human_selection_status=default_late | 结合人工说明和 role_in_group 确定层级；若明确为默认副引用，映射 secondary + yes；后位不等于 backup |
| human_order 与 proposed_order/legacy 字段不一致 | 已确认的人工顺序优先，原值保留；不要按 main/secondary 重新排序 |
| citation_group_id + citation_scope_id | 核对真实落点后确定 insert_as_group；一个主题可拆成多个括号，一个 scope 也可能包含多个位置 |
| tutorial_text_anchor=TP011 | 这是项目段落标识，需读取其映射并核对目标 Word，不直接作为搜索原文 |
| exact_tutorial_text 含竖线拼接多句 | 结合 placement_note 确定准确插入片段，不能整段拼接后直接搜索 |
| human_reviewed 与 tutorial_approved=no 并存 | 先确认各字段是选文状态、正文批准还是历史标记；不机械判断成冲突或已获写入授权 |
| word_insertion_authorized=no | 保留为历史来源状态；当前用户明确执行请求可更新本次授权记录，但不得在仅整理任务中自行转为授权 |
| scope_note / support_summary / cannot_support | 保留支撑边界，不能因列名适配丢失，也不能据此擅自改正文 |

## 6. 校验与自动衔接

第一阶段结束前检查：

1. 每个默认文献 ID 唯一映射到正确题名/DOI，选文、顺序和范围与人工决定一致。
2. 每个 group 对应一个实际落点，组内 anchor、position、occurrence 一致；目标 Word 已核对，不能仅对参考文本定位。
3. main/secondary/default_insert 转换正确；backup/not_use/manual_gate 未混入执行集合。
4. 增量修改已识别原有字段和 bibliography，保留、删除、新增、替换范围明确。
5. 严重未解决问题已列明；缺 key 等可在授权的第二阶段补齐的事项明确列为 pending。

自动衔接规则：

- 用户要求完整执行 + 第一阶段无阻塞项：直接读取主 workflow，开始第二阶段。
- 只缺可通过正常导入或读取补齐的 Zotero 信息：标为 ready_for_zotero_preparation，按授权进入准备步骤；身份/key 校验通过后继续 Word。
- 用户只要求整理：标为 preparation_only，交付三份文件后结束。
- 严重信息缺失、冲突、定位歧义或紧急异常：标为 blocked，暂停受影响操作并集中询问，可继续独立只读检查。该规则适用于两个阶段中的任意时刻。

第二阶段必须再次检查源/工作文档是否变化、key 是否仍有效。文件或文献库已变化时重做受影响的检查，不无条件沿用旧交接结果。

## 7. 与现有脚本的实际接口

`build_payload.py` 读取规范化 task table 和 crosswalk；required columns 以当前脚本为准。`--style` 必填，新文档由用户指定；仅刷新不调用构建器，沿用已有文档样式。

构建器拒绝非 add 的 action，但不执行 keep/replace/delete；human_selection_status、授权记录和人工支撑范围由执行 AI 解释。插入脚本在写入前验证全部 anchor，多匹配需明确 occurrence 和 selection reason。预整理仍负责选择依据正确，不能仅填字段来绕过歧义检查。

从零插入可将已校验的 add 行交给 payload 构建器。增量任务必须先由执行 AI 准备按现有 citationID/复合定位操作的项目级步骤，生成 add 或 replace 需要的新字段数据；keep/delete 行不可直接交给通用插入脚本。replace 必须更新或明确替换原字段，不能仅在相同位置再加一组。生成的 citationID 应检查与保留字段不冲突。

默认 bibliography-mode=auto：插入脚本复用已有一份 bibliography，仅在缺失时新增；omit 不新增。payload 绑定 docx 哈希，重复 citationID 或文档变化会阻止盲目重跑。更完整的增量衔接和验证见主 workflow 阶段 D-G。

若当前任务只是刷新验证，复用已有输入并跳过不需要的转换；若没有外部引用任务，不为满足脚本非空约束伪造引用。

## 8. 可复用启动语句

完整执行示例：

> 请使用 $zotero-word-autocite。结合本对话和项目文件，按 Skill 的输入说明整理任务表与文献映射，再按技术工作流完成已授权的 Zotero 准备及 Word 引用操作。目标 Word 和引用样式以我指定的内容为准。输入检查通过后自动衔接；严重缺失、冲突或异常时暂停询问。

仅预整理示例：

> 请按 PREPARE_CITATION_INPUTS.md 整理本项目已有引用方案，输出任务表、文献对照表和交接说明。本轮仅完成第一阶段。
