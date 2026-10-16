# Zotero Word AutoCite

**English** | [简体中文](README.zh-CN.md)

A Codex Skill for applying your citation plan in Microsoft Word with refreshable Zotero fields.

During manuscript preparation, references may be spread across PDF folders, Zotero collections, and tables, while citation locations are scattered through Word. This workflow brings those inputs together, matches the references you have selected to the intended locations, inserts or adjusts Zotero citation fields, then refreshes and checks the bibliography.

The result remains editable through Zotero's Word add-in. These are live citation fields, rather than manually typed reference numbers.

The homepage is bilingual. The Skill and detailed workflow documents are written in English.

## Choose the task you need

The workflow has several entry points. You do not need to run every stage for every request.

| Task | What you provide | Expected result |
|---|---|---|
| Read-only check | A question or files to examine | Findings or an explanation |
| Prepare citation inputs | Your decisions, notes, tables, or Word comments | A citation task table, reference mapping, and handoff notes |
| Prepare Zotero | Verified reference records, PDFs, or an existing collection | An import plan or prepared items and checked mappings, as authorized |
| Insert citations into a new document | Word file, citation plan, sources, and citation style | A working copy with Zotero fields, a bibliography if requested, and verification notes |
| Adjust existing citations | A Word file with Zotero fields and specified changes | Targeted changes, with unaffected citations preserved |
| Refresh and verify | A Word file with existing Zotero fields | A refreshed working copy and checks, using the existing style |

Citation plans can come from Markdown, CSV, JSON, Excel, Word comments, or ordinary written instructions. Sources can be existing Zotero items or PDFs with metadata. Codex checks what is already available before asking you to standardize it.

You decide which papers support which statements, their order, and where they belong. If that information is missing, the workflow can organize the inputs and identify the gaps; it does not choose the final citations on your behalf.

## What to prepare

- Windows, desktop Microsoft Word, Zotero, and the Zotero Word add-in.
- Python 3 and PowerShell.
- The target Word document, reference sources, and the citation plan or changes you want.
- A citation style for a new document. A refresh-only task can keep the document's existing style.

Before field operations, the workflow checks the actual local environment, including Zotero's local API and the official Word integration DLL. The Skill does not install or configure these applications for you.

## Installation and first use

The repository root is the Skill folder: `SKILL.md` sits beside `workflow/`. Keep the entire structure together.

You can clone it into the personal Codex Skill directory from PowerShell:

```powershell
git clone https://github.com/LiX-Works/zotero-word-autocite.git (Join-Path $env:USERPROFILE '.codex\skills\zotero-word-autocite')
```

Alternatively, download the repository and place the complete folder there. If the destination already exists, compare your customizations before replacing it. If Codex does not show the new Skill, reopen the task and check again.

For a new document, start with a request such as:

```text
Use $zotero-word-autocite. The project is at C:\path\to\project,
and the target file is manuscript.docx.

Use the project's citation plan and reference sources to insert Zotero
citation fields at the specified locations. Use the citation style I have
selected, then refresh and check the bibliography.
```

For an existing document, describe the affected passage and the exact reference changes. For example, identify one citation group using its surrounding sentence and the papers in your approved reference table, and ask to keep the rest of the document unchanged. A refresh-only request does not need a new citation plan or complete mapping tables.

## How changes and results are handled

For editing tasks, the workflow normally preserves the source document and works on a copy in your project. A clearly designated editable copy can be reused, with a backup before further changes. Inputs and reference identities are checked before writing.

When you have requested complete execution and the checks pass, Codex continues through the relevant stages under the existing authorization. It pauses the affected operation if a reference identity or location cannot be determined, information seriously conflicts, the document is damaged, or the save state is uncertain. Read-only or planning requests stop at their requested deliverable.

The batch scripts `build_payload.py` and `insert_zotero_fields.ps1` handle **new citations**. Replacing or deleting existing groups follows the workflow's targeted field-editing steps; those actions are not handled by simply appending new fields. Other citation groups are preserved.

PDF inventory and import-plan scripts prepare files and mappings. They do not perform the actual Zotero import themselves. Importing or changing Zotero records follows the task authorization and includes a database backup; the workflow does not default to deleting library items, PDFs, or all Word citations.

Completion means checking the actual fields, cited items, bibliography, and saved document. A successful script exit or a wait ending does not by itself prove that Zotero refresh completed. The handoff includes the working copy and a short account of what was verified or remains unfinished.

## Documentation and offline tests

| File or folder | Purpose |
|---|---|
| [SKILL.md](SKILL.md) | Task instructions and entry point |
| [Workflow overview](workflow/README_FIRST.md) | Choose between checking, preparation, insertion, adjustment, and refresh |
| [Input preparation](workflow/PREPARE_CITATION_INPUTS.md) | Interpret existing materials and prepare citation tasks and mappings |
| [Word–Zotero workflow](workflow/ZOTERO_WORD_AUTOMATION_WORKFLOW.md) | Field operations, refresh, troubleshooting, and verification |
| [Scripts](workflow/zotero_word_automation/scripts/) and [templates](workflow/zotero_word_automation/templates/) | Supporting operations and example input formats |
| [Tests](workflow/zotero_word_automation/tests/) | Offline regression tests |

To run the offline tests from the repository root:

```powershell
python -m unittest discover -s workflow/zotero_word_automation/tests -v
```

These tests do not launch Word or Zotero and do not prove that live insertion or refresh works in a particular desktop environment.

The project uses the [MIT license](LICENSE).
