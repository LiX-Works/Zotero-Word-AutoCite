# Zotero-Word Automated Citation Insertion Workflow

Version: 2026-09-16 (review revision)

Purpose: Guide Codex in batch-inserting refreshable Zotero citation fields into Word documents and generating/refreshing bibliographies on Windows + Microsoft Word + Zotero.
Core principle: Inspect inputs and context first, then produce normalized intermediate files. When information is insufficient, ask the user or list what is missing; do not guess.

Entry point: [README_FIRST.md](README_FIRST.md). For a Zotero-Word automation task, read the Skill entry point first, then the relevant sections of this workflow for the task type.

Two-phase interface: In Phase 1, follow `PREPARE_CITATION_INPUTS.md` in this directory to adapt existing project formats into a normalized task table, reference crosswalk, and handoff notes. In Phase 2, follow this file for Zotero preparation, Word citation operations, and verification. Stages A-G here are internal execution steps, not additional human confirmation gates. When the user has requested complete execution and input checks pass, continue automatically in the same AI conversation. For preparation-only requests, deliver Phase 1 and stop.

In Phase 2, read the exact input paths specified in the handoff, and check the Word document version and current Zotero state. Do not skip checks solely because an old handoff says ready. Phase 1 may leave keys/item IDs pending, but verification must be complete before writing to Word.

## 1. What this workflow does

The workflow uses the following operations rather than clicking through the Zotero add-in interface in Word:

1. Read the references to insert from the task table and Zotero library.
2. Obtain CSL JSON through Zotero's local API.
3. Generate a Word-Zotero field payload.
4. Write `ZOTERO_ITEM` and `ZOTERO_BIBL` fields through Zotero's official Word DLL.
5. Reopen Word and run `ZoteroRefresh`.
6. Check fields, numbering, bibliography, and rendered pages.

The inserted citations are refreshable Zotero fields, rather than plain-text numbers. After citations are added or removed in Word, Zotero can renumber them and update the bibliography.

## 2. Context checks before execution

Use README_FIRST's task routing to determine the applicable scope. The checks below apply to actual citation insertion/adjustment. For read-only tasks, inspect the relevant issue and reply without creating audit/normalized files unless requested. Refresh-only tasks retain existing document style and fields and do not require a complete task table/crosswalk; obtain clarification only if the style is missing or the user requests a change. Zotero-only preparation does not require Word anchors.

Before actual insertion/adjustment, check the applicable items below using the conversation, project files, and existing metadata. Reuse valid existing results without regenerating reports.

Establish:

1. Whether the user has specified the target Word document.
2. Whether source and project working-copy paths are explicit; for editing tasks, copy the source into the project folder and work on the copy by default.
3. Whether a citation insertion task table exists.
4. Whether the task table has sufficient fields.
5. Whether a unified reference crosswalk exists.
6. Whether the target references already exist in the Zotero library / collection.
7. If Zotero is not ready, whether PDF folders and metadata are available for importing or building the crosswalk.
8. Whether a new document has a specified citation style, or an existing document's original style can be verified and retained.
9. Whether Word or Zotero modifications are authorized, or only a plan is requested.
10. Whether this task is insertion from scratch, incremental changes, refresh and verification only, or planning only.
11. Whether a reference Word document contains comments or tracked changes, and whether it is only a reference rather than the final target.

If critical information is missing, prefer one of these approaches:

- If it can be reliably inferred and corrected from existing files, create normalized files without overwriting originals.
- Otherwise, ask the user or provide an explicit list of missing information.

Do not proceed with actual insertion when:

- The target Word document is unknown.
- The citations selected for default insertion are unknown.
- citation_id cannot be mapped to a Zotero key.
- The Zotero API cannot read the target item.
- A citation anchor has multiple matches without distinguishing context.
- Reference metadata is clearly wrong and the correction approach is unconfirmed.
- Document damage, an uncertain save state, or another serious exception affecting the result occurs during execution.

At the input stage or any later point, critical missing information, conflicts, or serious exceptions require pausing affected writes and asking the user. Explain completed operations, the issue and impact, and the decisions needed. Independent read-only checks may continue; do not merely record a warning and continue affected operations.

Consolidate multiple questions. Resolve low-risk issues independently when reliable, recording them in normalized files or logs. Do not create a separate dry-run report or fixed confirmation gate by default; proceed under existing authorization once checks pass.

## 3. Required inputs

### 3.1 Target Word document

Requirements:

- The file must exist.
- When the user requests actual editing, copy the source docx into the project folder and edit that working copy by default, without a separate declaration or repeated confirmation.
- If the user has explicitly designated an editable project copy, use it directly rather than copying repeatedly; automatically back it up before editing it again.
- Use a nonconflicting filename when creating a copy and retain the source. If it is unclear whether an existing project file is the source, create a new copy without overwriting it.
- Prefer closing any same-named document already open in Word.
- Do not edit multiple Word files simultaneously unless the user explicitly requests it.

If a reference Word document contains comments or revisions, such as comments or tracked changes from Word/WPS review mode, extract comment anchors and revision content independently first. Confirm precise locations before body-text edits or Zotero citation operations. Do not treat a commented reference Word document as the final editing target unless the user explicitly identifies it as such.

Suggested output: A project working copy named `target_zotero_YYYYMMDD.docx`; back up an existing working copy to `metadata/word_edit_backups_YYYYMMDD/`. Record source and working paths in the log without separately requesting permission to copy or edit the working copy. Audit-only and planning-only tasks do not proceed to actual editing.

### 3.2 Citation insertion task table

Task tables from different projects may have different structures. Audit and normalize them first.

Follow `PREPARE_CITATION_INPUTS.md` for field adaptation, precedence of human decisions, and the three handoff outputs. Incremental tasks should also include action (add/keep/replace/delete) and existing-field location information. The executing AI interprets these fields; the general-purpose insertion script does not route operations automatically.

Recommended normalized fields:

- `insert_as_group`: Citation-group ID for references within one bracket.
- `group_order`: Group insertion order.
- `claim_id`: Optional, from claim mapping.
- `target_location_description`: Human-readable citation-location description.
- `anchor_text`: Original Word passage used by the script for location.
- `anchor_occurrence`: Defaults to `unique`; an explicit choice may use `first` / `last` / `after_last_in_document` / a positive integer.
- `anchor_selection_reason`: Required for multiple matches; record the human plan or verified context supporting the selected location. first in an old table is not sufficient evidence for selecting among matches.
- `position`: `after` / `before`.
- `citation_id`: Project reference ID.
- `citation_order`: Reference order within the same bracket.
- `status`: `main` / `secondary` / `backup` / `not_use` / `manual_gate`.
- `default_insert`: `yes` / `no`.
- `notes`: Notes.

Execution rules:

- By default, insert only rows with `default_insert=yes` and status `main` or `secondary`.
- Do not insert `backup`, `not_use`, or `manual_gate` by default.
- Zotero refresh changes reference numbering. Deleting, moving, or replacing citations must never rely solely on current numeric labels in body text, such as 22, 23, or 24. Map them to title, DOI, citation_id, Zotero key, or an explicit task-table row.
- If task-table column names differ, create a normalized version first, for example:
  `metadata/word_citation_task_table_normalized_YYYYMMDD.csv`
- If `anchor_text` is missing, Codex may derive it from an explicit `target_location_description`, final plan, and the original Word text. If it agrees with the user's plan and identifies a unique location, write it into the normalized task table and continue. If multiple reasonable locations remain or the scope of support is unclear, pause that item and ask the user.

### 3.3 Unified reference crosswalk

Recommended fields:

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

Requirements:

- `citation_id` must map uniquely to a reference.
- `zotero_key` and `zotero_item_id` must be present for insertion into Word.
- Each key must return correct CSL JSON through Zotero's local API.
- `build_payload.py` compares the CSL URI key, normalized title, and nonempty DOI, and stops on mismatch. Title normalization ignores only HTML, case, whitespace, and punctuation; it does not perform fuzzy identity matching. The AI may investigate legitimate title variants and correct the normalized crosswalk. Do not bypass conflicts by loosening matching.
- Zotero preparation remains responsible for verifying the provenance of numeric `zotero_item_id` values. The CSL interface returns URIs, which alone do not prove the crosswalk's internal numeric IDs are correct. Retain the retrieval method and evidence that keys and IDs belong together in the same library.
- If no crosswalk exists, build one first from PDF folders, metadata, and Zotero import results.

### 3.4 PDF folders and Zotero import

Other projects may initially have only one or more PDF folders, without a Zotero library or crosswalk. This workflow covers that situation too.

Recommended steps:

1. Scan the PDF folders and produce a PDF inventory.
2. Merge with existing metadata and build an import manifest.
3. Import into Zotero or inspect an existing collection.
4. Export or generate a `citation_id -> zotero_key -> zotero_item_id` crosswalk.
5. Verify each reference through Zotero's local API.

Notes:

- Do not delete Zotero items or clear the library unless the user explicitly authorizes it.
- Back up `zotero.sqlite` before importing.
- Prefer linked PDF attachments to avoid copying or moving project PDFs.
- PDF filenames do not replace metadata. Verify DOI, title, authors, and journal separately.
- If Zotero's automatic PDF recognition is clearly wrong, stop and produce a list for human confirmation.

### 3.5 Zotero environment

Verify that:

- Zotero is running.
- Zotero's local API is reachable at `http://127.0.0.1:23119`.
- The Word add-in is installed.
- Zotero's official Word DLL can be found. Preferred path:
  `C:\Program Files\Zotero\integration\word-for-windows\libzoteroWinWordIntegration.dll`

By default, `insert_zotero_fields.ps1` searches common Zotero installation paths, including `C:\Program Files`, `C:\Program Files (x86)`, the user's local installation directory, running Zotero, `PATH`, and uninstall registry entries. For a custom installation, pass the full DLL path explicitly with `-ZoteroWordDllPath`. If it is still missing, the script must stop and report the searched locations. Codex must not guess a path; ask the user or list the missing requirement.

### 3.6 Citation style

Do not guess the submission style for a new document. For refresh-only or incremental changes that retain an existing style, read and reuse the existing DocumentData without asking again. Ask if the style is missing, conflicts, or the intended replacement is unclear. `build_payload.py` requires `--style`, which must explicitly receive the established style.

This project previously used Nano Letters:

`http://www.zotero.org/styles/nano-letters`

## 4. Standard execution stages

### Choose the execution mode

Insertion-from-scratch mode:

- Applies to a Word document without Zotero fields or one the user explicitly requests to rebuild.
- Requires a complete task table and crosswalk.
- Requires establishing whether to insert a bibliography.
- After required checks pass, write to the project working copy under the user's existing authorization.

Incremental-change mode:

- Applies to Word documents with existing Zotero fields.
- Preserve correct existing citations by default.
- Delete, replace, or add only the citation groups explicitly listed in the task table or user's change plan.
- Before editing, count existing `ZOTERO_ITEM`, `ZOTERO_BIBL`, and References entries.
- Produce keep / delete / add lists.
- Do not clear and rebuild all citation fields for convenience.

### Stage A: Input audit

Suggested outputs:

- `metadata/zotero_word_input_audit_YYYYMMDD.md`
- `metadata/zotero_word_input_audit_YYYYMMDD.csv`

Check:

- Whether the target Word document exists.
- Whether the task table exists.
- Whether the task table can be normalized.
- The number of task-table rows selected for default insertion.
- Whether every citation_id is in the crosswalk.
- Whether every Zotero key is readable through the API.
- Whether Word anchors exist and are unique.
- Whether manual gate / backup / not_use items were incorrectly included.
- Current counts of `ZOTERO_ITEM`, `ZOTERO_BIBL` fields and References entries.
- Whether this run is insertion from scratch or incremental changes.

### Stage B: Zotero preparation

With an existing collection:

1. Read the crosswalk.
2. Verify keys.
3. Verify title / DOI.
4. Correct metadata if needed, recording the changes.

With PDF folders only:

1. Scan PDFs.
2. Build an import manifest.
3. Check whether existing authorization includes importing into Zotero; continue if authorized and ask only if scope is unclear.
4. Back up the Zotero database.
5. Import items and linked attachments.
6. Generate the crosswalk.
7. Verify through the API.

### Stage C: Pre-execution validation

This is an internal check. Include results in the normalized task table, input audit, or execution log, rather than generating a separate dry-run report or stopping for confirmation by default.

At minimum, validate:

- Which Word document will be edited.
- Source and project working-copy paths, and the backup path for an existing copy, if applicable.
- Which citation groups will be inserted / deleted / replaced.
- The anchor near which each group will be inserted.
- Whether each anchor is unique.
- The citation_id / title / DOI / Zotero key members of each group.
- Which old references will disappear from References.
- Which new references will enter References.
- Whether all Zotero keys are readable through the API.
- Whether any warning / missing / ambiguous items remain.

Once checks pass, generate the payload and execute under existing authorization. Record and continue past nonblocking issues such as ordinary format differences or fields that can be completed reliably. Pause affected operations and consolidate questions for critical gaps, conflicts, ambiguous locations, or exceptions; this rule also applies to all later stages. If the user explicitly requests only a preview or plan, deliver that and stop.

### Stage D: Generate the payload

The general-purpose builder and insertion script handle only add, not group-level replace/delete. The builder rejects non-add actions to avoid reinserting keep rows and similar records. Older tables without action are treated as add.

Incremental handoff: First inventory citationIDs, member URIs/keys, source context, and the bibliography. Use project-specific steps to identify replacement/deletion fields uniquely from that information, then update the complete target group or delete the specified field while retaining other groups. Use a derived add-only table for additions. If project-specific steps implement replace by deleting then adding, record the complete target membership and order, and generate a payload bound to the current document hash after deletion. Do not merely rename replace as add and append it.

The payload includes the target docx SHA256. The insertion script stops before writing if the target has changed or a citationID already exists or is duplicated. Before a rerun, read the current state, mark completed groups keep, and generate a new payload only for remaining additions. Do not claim the general-purpose scripts provide fully idempotent incremental editing.

By default, `--bibliography-mode auto` adds a bibliography only if none exists and reuses one existing bibliography; pause to inspect multiple bibliographies. `omit` adds none. When citations already exist, the insertion script preserves DocumentData. The executing AI must first verify consistency with the target style; change styles using a separate document-style operation rather than implicitly overwriting it while adding fields.

The payload contains at least:

- `doc_path`
- `style`
- `doc_prefs`
- `groups`
- `bibliography`

Each group contains:

- `group_id`
- `anchor_text`
- `position`
- `citation_ids`
- `zotero_keys`
- `field_code`
- `placeholder`

Field-code format:

```text
ZOTERO_ITEM CSL_CITATION {...}
```

Bibliography field:

```text
ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY
```

### Stage E: Write Word fields

Reliable procedure:

1. Open the Word document.
   The general-purpose insertion script automatically backs up the target and checks the payload hash, all locations, duplicate citationIDs, and bibliography count before the first write. Multiple matches require an explicit occurrence and selection reason; the default is unique. It locates only in the main body; use project-specific steps for footnotes and other special locations.
2. Perform additions, replacements, or deletions according to the validated task table and user's change plan.
3. Optionally perform confirmed body-text replacements.
4. Locate anchors one by one in group order.
5. At each insertion point, call the DLL:
   - `insertField`
   - `setCode`
   - `setText`
6. Insert or reuse the bibliography field as needed.
7. Save and close Word.

Notes:

- Fields are inserted individually.
- ZoteroRefresh does not run after each field.
- Refresh once after all fields have been inserted.
- Clearing old Zotero fields is prohibited by default. It is allowed only when the user explicitly requests to "rebuild all citation fields," a complete task table and crosswalk exist, Zotero key verification passes, the original file or a pre-edit backup is retained, and pre-execution validation passes. Normally use incremental changes to existing Zotero fields rather than clearing and rebuilding.

### Stage F: Run ZoteroRefresh

Reliable procedure:

1. Reopen Word.
2. Make Word visible.
3. Run `ZoteroRefresh`.
4. Retain the previously validated wait of at least 120 seconds; then check that field codes and displayed results remain stable for at least 15 consecutive seconds, citations contain Zotero formatting information, and no empty displays or known placeholders remain. The default maximum observation time is 300 seconds.
5. After observation checks pass, save and check Word's Saved state.
6. Close Word.

Do not close Word immediately after `ZoteroRefresh` returns. Zotero may still be processing the document asynchronously.

These observations are supporting evidence, not a completion event provided by Zotero. Unchanged old fields can also meet the observation conditions, so Stage G must still verify the expected additions/deletions and error dialogs. `refresh_zotero_word.ps1 -PayloadPath ...` can additionally check this run's exact placeholders; the payload may be omitted when refreshing only an existing document.

On observation timeout, exceptions, or an uncertain save state, the script throws an error and leaves Word open without automatically saving/closing. The macro call or COM itself may block, and the script's MaxWaitSeconds does not cover blocking calls. The external runner must set an overall timeout and inspect logs/windows. A command timeout does not mean refresh is complete, and indiscriminately force-killing all Word processes is unacceptable. Apply the serious-exception rule before continuing.

### Stage G: Verification

Check:

- Number of `ZOTERO_ITEM` fields.
- Number of `ZOTERO_BIBL` fields.
- Whether plain-text numbers are mixed in beside Zotero fields.
- Whether placeholders remain.
- Whether the bibliography was generated.
- Whether citation numbers are visible.
- Whether expected removed references disappeared.
- Whether expected new references appeared.
- Whether grouped-citation and unique-bibliography-item counts are reasonable.
- Whether missing item / no longer exists errors appeared.
- Whether ZoteroRefresh completed and the document was saved successfully.
- Whether reference metadata is visibly correct.
- Whether Word opens normally.
- PDF/PNG rendering when needed, especially if page numbers, layout, or field visibility are uncertain.

`validate_docx_zotero.py` parses the main-body XML, reading field codes and displayed results separately. `--must-contain/--must-not-contain` inspect only visible body text; bibliography-specific options inspect only displayed bibliography-field results. A title in hidden JSON does not count as a visible reference match. The script checks complete field structure, duplicate citationIDs, empty displays, common placeholders, and exact payload placeholders; with a payload, it verifies group IDs and member URI order. Without a payload, it cannot identify all arbitrarily named placeholders.

For incremental verification, use `--payload-scope subset` and specify whole-document expected counts separately with `--expected-items/--expected-bibs`. Do not equate this run's added-group count with the whole-document field count. For bibliography additions/deletions, use `--bibliography-contains/--bibliography-not-contains` to check displayed text and manually verify unique items. Footnotes, headers, inherited hidden styles, layout, or adjacent plain-text numbers require separate checks.

Exit code 0 means only that static checks passed; the report always states `refresh_completion=not_determined_by_static_check`. A final success conclusion must combine refresh logs, save state, expected results, and actual Word display rather than relying only on an exit code or waiting duration.

Note: Task-table row count, grouped-citation count, Word field count, and References entry count need not match. One group may contain multiple references, while References counts only unique cited items.

Suggested outputs:

- `metadata/zotero_word_insert_log_YYYYMMDD.txt`
- `metadata/zotero_word_refresh_log_YYYYMMDD.txt`
- `metadata/zotero_word_validation_YYYYMMDD.json`
- `metadata/zotero_word_update_report_YYYYMMDD.md`

## 5. Common errors and handling

### 5.1 citation_id not found

Actions:

- Do not guess.
- Check aliases.
- Produce a missing list.
- Ask the user or build a new crosswalk.

### 5.2 Invalid Zotero key

Possible causes:

- The Zotero library was cleared and rebuilt.
- The item was deleted.
- The crosswalk is outdated.

Actions:

- Rematch by DOI/title in the current Zotero library.
- Generate a new key mapping.
- Do not reuse old keys without verification.

### 5.3 Incorrect Zotero metadata

Symptoms:

- Incorrect authors, titles, journals, or years in the bibliography.
- The PDF is correct, but recognition extracted irrelevant cover-page information.
- Only a note or attachment exists, without a parent item.

Actions:

- Correct the Zotero item using DOI / Crossref / the publisher page.
- Verify through the API again after correction.
- Run ZoteroRefresh again.

### 5.4 Nonunique anchor

Actions:

- Use longer context.
- Use start/end anchors.
- Use an occurrence index.
- Ask the user if it remains uncertain.

### 5.5 ZoteroRefresh error

Common dialogs:

- Zotero experienced an error updating the document.
- No running instance of Word was found.
- A word-processor integration command is already running.

Handling order:

1. Do not click repeatedly in quick succession.
2. Save what can currently be saved.
3. After confirming save state, close only this task's Word instance without affecting the user's other open documents.
4. Restart Zotero.
5. Reopen the document.
6. Run only ZoteroRefresh.
7. Allow sufficient time.
8. If it still fails, preserve logs and pause. Explain the failed stage, document save state, and decisions needed. If document damage or an uncertain save state has already been found, do not retry automatically; apply the serious-exception rule immediately.

### 5.6 Damaged Word field structure

Actions:

- Pause writing and inform the user, preserving the damaged copy and logs. Create a new recovery copy from the source or backup; do not automatically overwrite a file that may contain new user content.
- Do not repeatedly refresh damaged fields.
- Regenerate the payload.
- Reinsert the fields.

## 6. Conservative defaults for reuse

By default, do not:

- Delete the Zotero library.
- Delete project PDFs.
- Overwrite the original Word document.
- Insert backup/not_use/manual_gate items.
- Guess citation_id.
- Guess citation style.
- Treat plain-text numbers as success.
- Clear existing Zotero fields.
- Rely on current numeric Word labels to identify citations for deletion or replacement.

By default:

- Copy the source Word document into the project folder as a working copy; automatically back up an existing editable copy before reusing it.
- Back up the Zotero database for imports or metadata changes.
- Create normalized inputs as the task requires; refresh-only tasks do not need them.
- Create logs and a validation report for actual execution; report read-only findings in the reply by default.
- Ask the user when information is insufficient.

## 7. Recommended directory structure

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

Suggested project outputs:

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

The script package provides PDF inventory and Zotero import-plan generation scripts by default, not a general-purpose executor that directly modifies Zotero SQLite. Zotero database schemas, existing library state, duplicates, and merging rules vary by project. If automatic import is required, prepare project-specific import steps from `zotero_import_manifest`, back up first, and execute under existing authorization. Pause to ask only when the import strategy has a major ambiguity.

## 8. When to ask the user

Ask when:

- Multiple Word documents could be the target.
- The conversation does not establish whether actual edits or only an audit/plan are requested; when editing is already requested, do not ask again whether the working copy may be edited.
- The task table has no clear default insertion choices.
- Citation locations have multiple reasonable interpretations.
- Reference PDFs exist, but metadata is insufficient to create Zotero items.
- Duplicate Zotero items are found and it is unclear which to retain.
- A new document's style is unspecified, or an existing style cannot be established and this operation needs it.
- The Zotero library needs clearing or rebuilding.
- Existing Word citation fields need deleting without user authorization.
- Old Zotero fields need clearing without an explicit user request to rebuild.
- Any input or execution stage encounters critical gaps, conflicts, damaged documents/fields, accidental overwriting, an uncertain save state, unexpected library changes, or similar exceptions.

Suggested wording:

> The task table is normalized, but X citations each have two reasonable body-text locations that the existing plan cannot distinguish. I have paused writing those citations. Please confirm the target sentences in the list; independent read-only checks can continue.

## 9. Workflow boundaries

This workflow handles Word-Zotero citation automation, not whether a reference should be cited. Literature selection, claim mapping, and advisor-comment interpretation should be completed before entering this workflow or delivered as a separate preliminary plan.

Even if the user requests direct citation insertion, first audit the inputs. Execute only after the audit passes; do not skip it.
