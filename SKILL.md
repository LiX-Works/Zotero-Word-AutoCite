---
name: zotero-word-autocite
description: Organize citation inputs, prepare reference mappings and imports, insert or adjust refreshable Zotero citation fields, and refresh and verify bibliographies in Microsoft Word and Zotero on Windows. Use for Zotero-Word citation automation, additions, deletions, or changes to existing citations, citation-task adaptation, or related troubleshooting; literature discovery and scholarly reference selection are separate tasks.
---

# Zotero Word AutoCite

Read the [workflow entry point](workflow/README_FIRST.md) first, choose the route for the current task, and load only the relevant instructions.

## Inputs and execution

Determine the target Word document, citation plan, reference sources, style, and scope from the conversation and existing project files. Plans may come from Markdown, CSV, comments, or written instructions; do not require the user to prepare a standardized table first.

- For input adaptation, read the [input preparation guide](workflow/PREPARE_CITATION_INPUTS.md) and create or reuse the task table, crosswalk, and handoff notes.
- For Zotero preparation, field insertion or adjustment, refresh, or verification, read the relevant sections of the [technical workflow](workflow/ZOTERO_WORD_AUTOMATION_WORKFLOW.md).
- For a read-only check, read what the question requires and report the findings. Refreshing an existing document can retain its style without requiring a complete citation task table.
- When the user requests complete execution and checks pass, continue through the applicable stages under the existing authorization, without a separate dry-run confirmation gate.

At any stage, pause the affected operation and consolidate questions if critical information is missing, a conflict affects correctness, a reference identity or citation location cannot be established, the document is damaged, or its save state is uncertain. Resolve ordinary issues independently when they can be handled reliably.

Preserve human decisions about reference selection, order, and the scope of support. Follow the workflow's source-protection, backup, and identity-verification rules. Do not use current numeric Word citation labels as reference identities or insert backup, not_use, or manual_gate items by default.

## Tools and results

Scripts are in [zotero_word_automation/scripts](workflow/zotero_word_automation/scripts/) and templates in [templates](workflow/zotero_word_automation/templates/). Resolve these paths from this Skill directory; save task outputs in the user's project.

The execution environment is Windows, desktop Microsoft Word, Zotero, and its Word add-in. Verify the local API and DLL according to the technical workflow before actual field operations.

`build_payload.py` and `insert_zotero_fields.ps1` handle additions. For group replacement or deletion, use the project-specific steps in the technical workflow; do not append replace/delete actions as add operations. The Zotero import scripts generate plans; perform actual imports under the existing authorization through available interfaces.

On completion, verify the fields, actual cited members, bibliography, and saved result. A successful script exit or a wait ending does not by itself prove that ZoteroRefresh completed. Deliver the prepared inputs or the updated Word working copy and brief verification findings, as appropriate, and identify unfinished work.

`workflow/` is distributed with this Skill; script and template paths remain relative across projects. For Zotero or Word version differences, inspect the actual environment before execution.
