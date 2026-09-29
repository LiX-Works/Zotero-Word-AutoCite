# Zotero-Word 自动插入引用工作流

版本：2026-09-16（审查修订）

用途：指导 Codex 在 Windows + Microsoft Word + Zotero 环境中，批量向 Word 文档插入可刷新的 Zotero 引用字段，并生成/刷新参考文献表。
核心原则：先检查输入与上下文，再生成规范化中间文件；信息不足时询问用户或输出待补清单，不要猜测。

入口文件：[README_FIRST.md](README_FIRST.md)。开始 Zotero-Word 自动化任务时，Codex 先读 Skill 入口，再按任务类型阅读本 workflow 的相关章节。

两阶段接口：第一阶段按同目录 `PREPARE_CITATION_INPUTS.md` 将项目原有格式适配为规范化任务表、文献对照表和交接说明；第二阶段按本文件完成 Zotero 准备、Word 引用操作及验证。本文件的阶段 A-G 是内部执行步骤，不是额外人工确认关卡。用户已要求完整执行且输入检查通过时，同一 AI 对话应自动衔接；仅要求整理时交付第一阶段后停止。

第二阶段读取交接说明指定的准确输入路径，核对 Word 版本与当前 Zotero 状态；不得只凭旧交接文件中的 ready 状态跳过检查。第一阶段允许 key/item ID 待补，进入 Word 写入前必须完成验证。

## 1. 这个工作流解决什么问题

本工作流不是用鼠标点击 Word 里的 Zotero 插件界面，而是：

1. 从任务表和 Zotero 文献库读取需要插入的文献；
2. 用 Zotero 本地 API 取得 CSL JSON；
3. 生成 Word-Zotero 字段 payload；
4. 通过 Zotero 官方 Word DLL 写入 `ZOTERO_ITEM` 和 `ZOTERO_BIBL` 字段；
5. 重新打开 Word，运行 `ZoteroRefresh`；
6. 检查字段、编号、参考文献表和渲染页面。

这样插入的不是普通文本编号，而是 Zotero 可刷新字段。后续 Word 中增删引用后，Zotero 可以重新编号并更新 bibliography。

## 2. Codex 执行前必须先做的上下文检查

先按 README_FIRST 的任务路由判断适用范围。以下检查用于实际插入/调整引用；只读任务只检查相关问题并回复，除非要求输出文件，不创建 audit/normalized 等文件。仅刷新沿用文档现有样式和字段，不需要完整 task table/crosswalk；样式缺失或用户要求更换时才补充确认。仅 Zotero 准备不要求 Word anchor。

实际插入/调整前，根据对话、项目文件和已有 metadata 检查以下适用项；已有结果有效时复用，无需重复生成报告。

必须确认：

1. 用户是否指定了目标 Word 文档。
2. 源 Word 和项目内工作副本路径是否明确；修改任务默认复制到项目文件夹后操作副本。
3. 是否存在引用插入任务表。
4. 引用插入任务表是否含有足够字段。
5. 是否存在统一文献对照表。
6. Zotero library / collection 是否已有目标文献。
7. 如果 Zotero 未准备好，是否存在 PDF 文件夹和 metadata，可用于导入或建立对照表。
8. 新文档是否明确引用样式；已有文档是否可核实并沿用原样式。
9. 是否允许修改 Word、Zotero、或只允许生成计划。
10. 当前任务是从零插入、增量修改、只刷新验证，还是只生成计划。
11. 是否有参考 Word 含批注或修订痕迹，且它是否只是参考文件而非最终目标文档。

如果任何关键信息缺失，优先采取下面两种方式之一：

- 可以从现有文件可靠推断并自动修正时，新建规范化文件，不要覆盖原文件。
- 无法可靠推断时，向用户询问或输出明确的待补信息清单。

禁止在以下情况下继续正式插入：

- 不知道目标 Word 文档是哪一个；
- 不知道哪些 citation 默认插入；
- citation_id 无法映射到 Zotero key；
- Zotero API 无法读取目标条目；
- 引用落点 anchor 不唯一且无上下文区分；
- 文献元数据明显错误但未确认修正方式。
- 运行中出现文档损坏、无法确认是否保存或其他影响结果的严重异常。

输入阶段及任务进行的任意阶段，发现严重信息缺失、信息冲突或紧急异常时，必须暂停受影响的写入并向用户询问。说明已完成操作、问题及影响、需要用户决定的事项；可继续独立的只读检查，不得仅记录警告后继续受影响的操作。

多个疑问应合并询问。低风险且能可靠解决的问题可自动处理并记录在 normalized 文件或日志中。默认不生成独立 dry-run 报告、不设置固定确认关卡；检查通过后在已有授权范围内继续。

## 3. 前置输入

### 3.1 目标 Word 文档

要求：

- 文件必须存在；
- 用户要求实际修改时，默认把源 docx 复制到项目文件夹并直接修改该工作副本，无需额外声明或逐次确认；
- 若用户已明确指定项目内可修改副本，直接使用，不反复复制；再次修改已有工作副本前自动备份；
- 创建副本时使用不冲突的文件名，保留源文件。项目内文件是否为源文件不明确时，新建副本，不覆盖；
- 最好关闭 Word 中已打开的同名文档；
- 不要同时修改多个 Word 文件，除非用户明确要求。

如果参考 Word 中含有批注类内容，例如由 Word/WPS 修订模式得到的批注或修订痕迹，必须先独立提取批注锚点和修订内容，确认能精确定位后，再进入正文修改或 Zotero 引用操作流程。不要把带批注的参考 Word 当作最终待修改正文，除非用户明确说明它就是目标文档。

建议输出：项目内 `target_zotero_YYYYMMDD.docx` 工作副本；需要备份已有工作副本时，保存到 `metadata/word_edit_backups_YYYYMMDD/`。日志记录源路径和工作路径即可，无需单独请求复制或修改副本的确认。仅审计或计划任务不进入正式修改。

### 3.2 引用插入任务表

任务表可能来自不同项目，结构不一定相同。因此要先审计并规范化。

字段适配、人工决定优先级和三份交接输出遵循 `PREPARE_CITATION_INPUTS.md`。增量任务还应包含 action（add/keep/replace/delete）及已有字段定位信息；这些字段由执行 AI 解释，通用插入脚本不会自动分流。

推荐规范字段：

- `insert_as_group`：同一个括号内的引用分组 ID。
- `group_order`：插入组顺序。
- `claim_id`：可选，来自 claim mapping。
- `target_location_description`：给人看的落点描述。
- `anchor_text`：脚本用于定位的 Word 原文片段。
- `anchor_occurrence`：默认 `unique`；明确选择时可为 `first` / `last` / `after_last_in_document` / 正整数。
- `anchor_selection_reason`：多匹配时必填，记录选择某一处的人工方案或已核实上下文。旧表中的 first 不能单独作为多匹配选择依据。
- `position`：`after` / `before`。
- `citation_id`：项目内文献 ID。
- `citation_order`：同一括号内的文献顺序。
- `status`：`main` / `secondary` / `backup` / `not_use` / `manual_gate`。
- `default_insert`：`yes` / `no`。
- `notes`：备注。

执行规则：

- 默认只插入 `default_insert=yes` 且 status 为 `main` 或 `secondary` 的行。
- `backup`、`not_use`、`manual_gate` 不默认插入。
- Zotero 刷新后参考文献编号会变化。任何删除、移动、替换引用的操作，都不能只依赖当前正文中的数字编号，例如 22、23、24。必须映射到 title、DOI、citation_id、Zotero key 或明确的 task table 行。
- 如果任务表列名不同，Codex 应先生成 normalized 版本，例如：
  `metadata/word_citation_task_table_normalized_YYYYMMDD.csv`
- 如果 `anchor_text` 缺失，Codex 可根据明确的 `target_location_description`、final plan 和 Word 原文补齐；与用户方案一致且定位唯一时，写入规范化任务表后继续。存在多个合理落点或支撑范围不明时，暂停该项并询问用户。

### 3.3 统一文献对照表

推荐字段：

- `citation_id`
- `aliases`
- `title`
- `doi`
- `year`
- `first_author`
- `journal`
- `pdf_path`
- `zotero_key`
- `zotero_item_id`
- `zotero_collection`
- `source_metadata_file`
- `notes`

要求：

- `citation_id` 必须能唯一映射到文献。
- 执行 Word 插入时，`zotero_key` 和 `zotero_item_id` 必须存在。
- 每个 key 必须能通过 Zotero local API 返回正确 CSL JSON。
- `build_payload.py` 会比对 CSL URI 的 key、规范化后的题名，以及非空 DOI；不一致时停止。题名规范化仅忽略 HTML、大小写、空白和标点，不做模糊身份匹配。正常题名变体可由 AI 查清并修正规范化对照表；冲突不能靠放宽匹配绕过。
- 数字 `zotero_item_id` 的来源核验仍由 Zotero 准备步骤承担；CSL 接口返回 URI，不能仅凭此证明对照表中的内部数字 ID 正确。保留获取方式和同库 key/ID 对应证据。
- 如果 crosswalk 不存在，必须先根据 PDF 文件夹、metadata、Zotero 导入结果建立。

### 3.4 PDF 文件夹与 Zotero 导入

其他项目正常情况下可能只有一个或多个 PDF 文件夹，还没有 Zotero library 或 crosswalk。这种情况也纳入本工作流。

推荐步骤：

1. 扫描 PDF 文件夹，生成 PDF inventory。
2. 与已有 metadata 合并，建立 import manifest。
3. 导入 Zotero 或检查已有 collection。
4. 导出或生成 `citation_id -> zotero_key -> zotero_item_id` 对照表。
5. 用 Zotero local API 验证每条文献。

注意：

- 不要删除 Zotero 条目或清空 library，除非用户明确授权。
- 导入 Zotero 前备份 `zotero.sqlite`。
- 尽量使用 linked PDF attachment，避免复制或移动项目 PDF。
- PDF 文件名不能替代元数据。DOI、题名、作者、期刊必须另行核对。
- 如果 Zotero 自动识别 PDF 结果明显错误，应停止并生成待人工确认清单。

### 3.5 Zotero 环境

必须确认：

- Zotero 正在运行；
- Zotero local API 可访问：`http://127.0.0.1:23119`;
- Word 插件已安装；
- Zotero 官方 Word DLL 可找到。优先路径为：
  `C:\Program Files\Zotero\integration\word-for-windows\libzoteroWinWordIntegration.dll`

`insert_zotero_fields.ps1` 默认自动搜索常见 Zotero 安装路径，包括 `C:\Program Files`、`C:\Program Files (x86)`、用户本地安装目录、正在运行的 Zotero、`PATH` 和卸载注册表。自定义安装可用 `-ZoteroWordDllPath` 显式传入 DLL 完整路径。仍找不到时，脚本必须停止并输出已检查范围；Codex 不得猜测路径，应向用户询问或输出待补清单。

### 3.6 引用样式

新文档不猜测投稿样式。已有文档仅刷新或保留原样式的增量修改，可读取已有 DocumentData 后沿用，不重复询问；样式缺失、冲突或更换目标不明时询问。`build_payload.py` 的 `--style` 为必填，需显式传入确定的样式。

本项目曾使用 Nano Letters：

`http://www.zotero.org/styles/nano-letters`

## 4. 标准执行阶段

### 执行模式选择

从零插入模式：

- 适用于无 Zotero 字段或用户明确要求重建的 Word；
- 需要完整 task table 和完整 crosswalk；
- 需要确认是否插入 bibliography；
- 必要检查通过后，按用户已有授权写入项目内工作副本。

增量修改模式：

- 适用于已有 Zotero 字段的 Word；
- 默认保留已有正确引用；
- 只删除、替换、添加任务表或用户修改方案明确列出的引用组；
- 修改前必须统计已有 `ZOTERO_ITEM`、`ZOTERO_BIBL` 和 References 条目；
- 输出保留 / 删除 / 新增清单；
- 不得为了省事清空重建全部引用字段。

### 阶段 A：输入审计

输出建议：

- `metadata/zotero_word_input_audit_YYYYMMDD.md`
- `metadata/zotero_word_input_audit_YYYYMMDD.csv`

检查项：

- 目标 Word 是否存在；
- 任务表是否存在；
- 任务表是否可规范化；
- 任务表默认插入行数；
- citation_id 是否都能在 crosswalk 中找到；
- Zotero key 是否都能 API 读取；
- Word anchor 是否存在且唯一；
- 是否有 manual gate / backup / not_use 被错误纳入。
- 当前 Word 中已有多少 `ZOTERO_ITEM`、`ZOTERO_BIBL` 字段和 References 条目；
- 本轮属于从零插入模式还是增量修改模式。

### 阶段 B：Zotero 准备

有现成 collection 时：

1. 读取 crosswalk；
2. 验证 key；
3. 验证 title / DOI；
4. 必要时修正元数据，但必须记录。

只有 PDF 文件夹时：

1. 扫描 PDF；
2. 建 import manifest；
3. 核对已有授权是否包含导入 Zotero；已授权则继续，范围不明确才询问；
4. 备份 Zotero 数据库；
5. 导入条目和 linked attachment；
6. 生成 crosswalk；
7. API 验证。

### 阶段 C：执行前校验

本阶段是内部检查，结果并入规范化任务表、输入审计或执行日志，不单独生成 dry-run 报告，不默认停下来等待用户确认。

校验至少覆盖：

- 将修改哪个 Word；
- 源文档和项目内工作副本路径，以及已有副本的备份路径（如适用）；
- 将插入 / 删除 / 替换哪些 citation group；
- 每组引用插在哪个 anchor 附近；
- anchor 是否唯一；
- 每组包含哪些 citation_id / title / DOI / Zotero key；
- 哪些旧文献会从 References 中消失；
- 哪些新文献会加入 References；
- Zotero key 是否都能被 API 读取；
- 是否存在 warning / missing / ambiguous 项。

检查通过后在已有授权范围内直接生成 payload 并执行。普通格式差异、可可靠补齐的字段等非阻塞问题，记录后继续。严重缺失、冲突、落点歧义或异常则暂停受影响的操作并集中询问；此规则同样适用于后续所有阶段。用户明确要求只做预览或计划时，按要求输出后停止。

### 阶段 D：生成 payload

通用构建器和插入脚本只负责 add，不实现按组 replace/delete。构建器遇到非 add 的 action 会拒绝，避免把 keep 等行再次插入；旧表无 action 时按 add 处理。

增量修改衔接：先盘点 citationID、成员 URI/key、原文上下文和 bibliography；项目级步骤按这些信息唯一定位要替换/删除的字段，再更新完整目标组或删除指定字段，保留其他组。新增部分使用派生的 add-only 表。若项目步骤将 replace 转成删除后新增，必须记录完整目标成员和顺序，删除后再生成绑定当前文档哈希的 payload。不能只是把 replace 改名为 add 后追加。

payload 带目标 docx 的 SHA256；插入脚本发现目标已变化、citationID 已存在或重复时，在写入前停止。再次运行必须读取现状，已完成组标为 keep，只对剩余新增组生成新的 payload；不宣称通用脚本实现了完整幂等增量编辑。

`--bibliography-mode auto` 默认仅在不存在 bibliography 时新增，已有一份则复用；多份时暂停检查。`omit` 不新增 bibliography。已有引用时插入脚本保留 DocumentData；执行 AI 须先核对目标样式一致，需换样式则使用单独的文档样式操作，不在新增字段时隐式覆盖。

payload 至少包含：

- `doc_path`
- `style`
- `doc_prefs`
- `groups`
- `bibliography`

每个 group 包含：

- `group_id`
- `anchor_text`
- `position`
- `citation_ids`
- `zotero_keys`
- `field_code`
- `placeholder`

字段代码格式：

```text
ZOTERO_ITEM CSL_CITATION {...}
```

bibliography 字段：

```text
ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY
```

### 阶段 E：写入 Word 字段

稳定做法：

1. 打开 Word 文档；
   通用插入脚本自动备份目标，校验 payload 哈希、所有落点、重复 citationID 和 bibliography 数量后才首次写入。多匹配需明确 occurrence 和 selection reason；缺省为 unique。它只在主正文中定位，脚注等特殊位置用项目级步骤。
2. 根据已校验的任务表和用户修改方案执行新增、替换或删除；
3. 可选：执行经过确认的正文替换；
4. 按 group 顺序逐个定位 anchor；
5. 每个插入点调用 DLL：
   - `insertField`
   - `setCode`
   - `setText`
6. 按需插入或复用 bibliography 字段；
7. 保存并关闭 Word。

注意：

- 字段是逐个插入的；
- ZoteroRefresh 不在每个字段后运行；
- 所有字段插完后统一刷新一次。
- 默认禁止清除旧 Zotero 字段。只有用户明确要求“重建全部引用字段”，并且已有完整 task table、完整 crosswalk、Zotero key 验证通过、原文件或修改前备份已保留、执行前校验通过后，才允许清除旧字段。一般情况下应采用“增量修改已有 Zotero 字段”，而不是清空重建。

### 阶段 F：运行 ZoteroRefresh

稳定做法：

1. 重新打开 Word；
2. Word 设为 visible；
3. 运行 `ZoteroRefresh`；
4. 保留已验证的至少 120 秒等待；之后检查字段代码和显示结果连续稳定至少 15 秒、引用含 Zotero 格式化信息、无空显示或已知占位符；默认最大观察时间 300 秒；
5. 观察通过后保存，并检查 Word 的 Saved 状态；
6. 关闭 Word。

不要在 `ZoteroRefresh` 返回后立刻关闭 Word。Zotero 后台可能仍在异步处理文档。

这些观察是辅助证据，不是 Zotero 提供的完成事件。未变化的旧字段也可能满足观察条件，所以仍需阶段 G 核验本次期待的增删结果和错误弹窗。`refresh_zotero_word.ps1 -PayloadPath ...` 可额外检查本次精确占位符；仅刷新旧文档时可不提供 payload。

脚本在观察超时、异常或保存状态不明时抛错并保留 Word 现场，不自动保存/关闭。宏调用或 COM 本身可能阻塞，脚本的 MaxWaitSeconds 不覆盖阻塞调用；外部执行器须设置总超时并查看日志/窗口，不能把命令超时等同于刷新完成，也不能盲目强杀所有 Word。继续前按严重异常原则处理。

### 阶段 G：验证

必须检查：

- `ZOTERO_ITEM` 字段数量；
- `ZOTERO_BIBL` 字段数量；
- 是否有普通文本编号混在 Zotero 字段旁边；
- 是否有占位符残留；
- bibliography 是否生成；
- 引用编号是否可见；
- expected removed references 是否消失；
- expected new references 是否出现；
- grouped citation 数和 unique bibliography item 数是否合理；
- 是否出现 missing item / no longer exists 类错误；
- ZoteroRefresh 是否完成并保存成功；
- 文献元数据是否明显正确；
- Word 是否可正常打开；
- 必要时检查 PDF/PNG 渲染，尤其是页码、排版或字段可见性存在疑点时。

`validate_docx_zotero.py` 解析主正文 XML，分别读取字段代码和显示结果；`--must-contain/--must-not-contain` 只检查正文可见文本，bibliography 专用选项只检查参考文献字段的显示结果。隐藏 JSON 中的题名不作为可见参考文献命中。它检查完整字段结构、重复 citationID、空显示、常见及 payload 精确占位符；带 payload 时核验组 ID 与成员 URI 顺序。没有 payload 时无法识别所有任意命名占位符。

增量验证用 `--payload-scope subset`，另以 `--expected-items/--expected-bibs` 给出全篇预期数量；不要用本轮新增组数等同于全文字段数。参考文献增删用 `--bibliography-contains/--bibliography-not-contains` 检查显示文本，并人工核对 unique items；脚注、页眉、继承隐藏样式、版面或相邻普通文本编号需另查。

脚本返回 0 只表示静态检查通过，报告固定标明 `refresh_completion=not_determined_by_static_check`。最终成功结论需结合刷新日志、保存状态、期望结果和实际 Word 显示，不能仅凭退出码或等待时长。

注意：task table 行数、grouped citation 数、Word 字段数、References 条目数不一定相等。原因是一个 group 可以包含多篇文献，而 References 只统计 unique cited items。

建议输出：

- `metadata/zotero_word_insert_log_YYYYMMDD.txt`
- `metadata/zotero_word_refresh_log_YYYYMMDD.txt`
- `metadata/zotero_word_validation_YYYYMMDD.json`
- `metadata/zotero_word_update_report_YYYYMMDD.md`

## 5. 常见错误与处理

### 5.1 citation_id 找不到

处理：

- 不要猜；
- 检查 aliases；
- 生成 missing list；
- 询问用户或建立新的 crosswalk。

### 5.2 Zotero key 失效

原因：

- Zotero library 被清空重建；
- 条目被删除；
- crosswalk 是旧的。

处理：

- 用 DOI/title 在当前 Zotero 中重新匹配；
- 生成新的 key mapping；
- 不要直接沿用旧 key。

### 5.3 Zotero 元数据错误

表现：

- bibliography 中作者、题名、期刊、年份异常；
- PDF 正确但条目识别成封面干扰信息；
- 只有笔记或附件，没有父条目。

处理：

- 用 DOI / Crossref / publisher 页面修正 Zotero 条目；
- 修正后重新 API 验证；
- 再运行 ZoteroRefresh。

### 5.4 anchor 不唯一

处理：

- 使用更长上下文；
- 使用 start/end anchor；
- 使用 occurrence index；
- 仍不确定时询问用户。

### 5.5 ZoteroRefresh 报错

常见弹窗：

- Zotero 在更新文档时出错；
- 没有找到运行中的 Word 程序；
- 文字处理器集成命令已在运行。

处理顺序：

1. 不要重复快速点击；
2. 保存当前可保存内容；
3. 确认保存状态后只关闭本任务 Word 实例，不影响用户其他已打开文档；
4. 重启 Zotero；
5. 重新打开文档；
6. 只运行 ZoteroRefresh；
7. 等待足够时间；
8. 若仍失败，保留日志并暂停，向用户说明失败阶段、文档保存状态和待决定事项。若已发现文档损坏或保存状态不明，不自动重试，立即按严重异常处理。

### 5.6 Word 字段结构损坏

处理：

- 暂停写入并告知用户，保留损坏副本及日志；从源文件或备份新建恢复副本，不自动覆盖可能包含用户新增内容的文件；
- 不要在损坏字段上反复刷新；
- 重新生成 payload；
- 重新插入字段。

## 6. 复用时的保守默认

默认不做：

- 不删除 Zotero library；
- 不删除项目 PDF；
- 不覆盖原始 Word；
- 不把 backup/not_use/manual_gate 插入；
- 不猜 citation_id；
- 不猜引用样式；
- 不把普通文本编号当成功。
- 不默认清除已有 Zotero 字段；
- 不依赖当前 Word 中的数字编号作为引用删除或替换依据。

默认要做：

- 将源 Word 复制到项目文件夹作为工作副本；复用已有可修改副本时自动备份；
- 备份 Zotero 数据库，若涉及导入或元数据修改；
- 按任务需要新建 normalized input；只刷新无需生成；
- 实际执行新建日志和 validation report；只读检查默认在回复中汇报；
- 信息不足时问用户。

## 7. 推荐目录结构

```text
zotero-word-autocite\
  SKILL.md
  workflow\
    README_FIRST.md
    PREPARE_CITATION_INPUTS.md
    ZOTERO_WORD_AUTOMATION_WORKFLOW.md
    zotero_word_automation\
      scripts\
        inventory_pdf_folders.py
        prepare_zotero_import_plan.py
        build_payload.py
        insert_zotero_fields.ps1
        refresh_zotero_word.ps1
        validate_docx_zotero.py
      templates\
        citation_task_table_template.csv
        citation_crosswalk_template.csv
        zotero_import_manifest_template.csv
```

项目内建议输出：

```text
metadata\
  zotero_word_input_audit_YYYYMMDD.md
  word_citation_task_table_normalized_YYYYMMDD.csv
  unified_citation_crosswalk_YYYYMMDD.csv
  citation_input_handoff_YYYYMMDD.md
  word_zotero_payload_YYYYMMDD.json
  word_zotero_insert_log_YYYYMMDD.txt
  word_zotero_refresh_log_YYYYMMDD.txt
  word_zotero_validation_YYYYMMDD.json
  word_zotero_update_report_YYYYMMDD.md
```

说明：本脚本包默认提供 PDF 清单与 Zotero 导入计划生成脚本，不提供直接改 Zotero SQLite 的通用执行脚本。Zotero 数据库 schema、已有库状态、重复条目及合并规则具有项目差异。若需要自动导入，Codex 应基于 `zotero_import_manifest` 准备项目级导入步骤，完成备份后在已有授权范围内执行；导入策略有重大歧义时才暂停询问。

## 8. 何时应该询问用户

必须询问用户的情况：

- 多个 Word 文档都可能是目标文档；
- 对话无法确定用户要求实际修改还是仅审计/计划；用户已要求修改时，不重复询问工作副本能否修改；
- 任务表没有明确默认插入项；
- 引用落点存在多种合理解释；
- 文献 PDF 有但 metadata 不足以建立 Zotero 条目；
- Zotero 中发现重复条目且无法判断保留哪一个；
- 新文档样式未指定，或已有样式无法确定且本次操作需要它；
- 需要清空或重建 Zotero library；
- 需要删除现有 Word 引用字段但用户未授权。
- 需要清除旧 Zotero 字段但用户未明确要求重建；
- 输入或执行任意阶段出现严重信息缺失、冲突、文档/字段损坏、意外覆盖、无法确认保存状态或 library 状态意外变化等异常。

建议话术：

> 已完成任务表规范化，但其中 X 个引用在正文中各有两个合理落点，现有方案无法区分。我已暂停这些引用的写入；请确认清单中的目标句子，其余只读检查可继续。

## 9. 本流程的边界

本工作流只负责 Word-Zotero 引用自动化，不负责判断文献是否该引用。文献筛选、claim mapping、导师批注处理应在进入本流程前完成，或作为前置阶段单独输出计划。

如果用户要求“直接插引用”，Codex 仍应先做输入审计。审计通过后再执行，不要跳过。
