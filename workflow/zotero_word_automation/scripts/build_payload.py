#!/usr/bin/env python3
"""Build a Word-Zotero insertion payload from a normalized task table.

This script intentionally assumes normalized CSV inputs. If a project's task
table has different columns, create a normalized copy first instead of making
this script guess.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import re
import sys
import unicodedata
import urllib.request
import uuid
from collections import defaultdict
from pathlib import Path
from xml.sax.saxutils import quoteattr


REQUIRED_TASK_COLUMNS = {
    "insert_as_group",
    "anchor_text",
    "position",
    "citation_id",
    "citation_order",
    "status",
    "default_insert",
}

REQUIRED_CROSSWALK_COLUMNS = {
    "citation_id",
    "title",
    "doi",
    "zotero_key",
    "zotero_item_id",
}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def require_columns(rows: list[dict[str, str]], required: set[str], label: str) -> None:
    if not rows:
        raise SystemExit(f"{label} is empty")
    columns = set(rows[0].keys())
    missing = sorted(required - columns)
    if missing:
        raise SystemExit(f"{label} missing required columns: {', '.join(missing)}")


def fetch_csl_json(zotero_key: str, timeout: int = 20) -> dict:
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    url = f"http://127.0.0.1:23119/api/users/0/items/{zotero_key}?format=csljson"
    with opener.open(url, timeout=timeout) as response:
        data = json.loads(response.read().decode("utf-8"))
    if not data:
        raise RuntimeError(f"Empty CSL response for Zotero key {zotero_key}")
    item = data[0]
    if "id" not in item:
        raise RuntimeError(f"CSL response for {zotero_key} does not include item URI id")
    return item


def yes(value: str) -> bool:
    return str(value or "").strip().lower() in {"yes", "y", "true", "1"}


def normalize_title(value: str) -> str:
    text = html.unescape(re.sub(r'<[^>]+>', '', value or ''))
    return ''.join(c for c in unicodedata.normalize('NFKC', text).casefold() if c.isalnum())


def normalize_doi(value: str) -> str:
    return re.sub(r'^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)', '', (value or '').strip(), flags=re.I).lower()


def verify_identity(row: dict, item: dict) -> None:
    cid = row['citation_id']
    if not str(item.get('id', '')).rstrip('/').endswith('/items/' + row['zotero_key'].strip()):
        raise SystemExit(f'{cid}: CSL URI does not match Zotero key')
    title = normalize_title(row.get('title', ''))
    if not title or title != normalize_title(item.get('title', '')):
        raise SystemExit(f'{cid}: title mismatch between crosswalk and Zotero; reconcile metadata before writing')
    doi = normalize_doi(row.get('doi', ''))
    if doi and doi != normalize_doi(item.get('DOI', '')):
        raise SystemExit(f'{cid}: DOI mismatch or missing Zotero DOI')


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True, help="Target Word .docx path")
    parser.add_argument("--task-table", required=True, help="Normalized citation task CSV")
    parser.add_argument("--crosswalk", required=True, help="Citation ID to Zotero key CSV")
    parser.add_argument("--output", required=True, help="Output payload JSON")
    parser.add_argument("--style", required=True, help="Explicit, verified Zotero style ID")
    parser.add_argument("--include-status", default="main,secondary")
    parser.add_argument("--bibliography-heading", default="References")
    parser.add_argument("--bibliography-mode", choices=["auto", "omit"], default="auto",
                        help="Auto inserts only if absent; omit preserves existing bibliography without adding one")
    args = parser.parse_args()

    docx = Path(args.docx)
    if not docx.exists():
        raise SystemExit(f"Target docx does not exist: {docx}")

    task_rows = read_csv(Path(args.task_table))
    crosswalk_rows = read_csv(Path(args.crosswalk))
    require_columns(task_rows, REQUIRED_TASK_COLUMNS, "task table")
    require_columns(crosswalk_rows, REQUIRED_CROSSWALK_COLUMNS, "crosswalk")

    unsupported = {r.get('action', '').strip() for r in task_rows} - {'', 'add'}
    if unsupported:
        raise SystemExit('Add-only builder cannot execute actions: ' + ', '.join(sorted(unsupported)) +
                         '. Use project-level incremental steps and an add-only derived table.')

    include_status = {x.strip() for x in args.include_status.split(",") if x.strip()}
    if not include_status <= {'main', 'secondary'}:
        raise SystemExit('Only main/secondary can be included by this writer')
    active_rows = [
        r for r in task_rows
        if yes(r.get("default_insert", "")) and (r.get("status", "").strip() in include_status)
    ]
    if not active_rows:
        raise SystemExit("No active insertion rows after default_insert/status filtering")

    by_cid: dict[str, dict[str, str]] = {}
    duplicate_cids: list[str] = []
    for row in crosswalk_rows:
        cid = row.get("citation_id", "").strip()
        if not cid:
            continue
        if cid in by_cid:
            duplicate_cids.append(cid)
        by_cid[cid] = row
    if duplicate_cids:
        raise SystemExit(f"Duplicate citation_id in crosswalk: {', '.join(sorted(set(duplicate_cids)))}")

    missing = sorted({r["citation_id"].strip() for r in active_rows if r["citation_id"].strip() not in by_cid})
    if missing:
        raise SystemExit(f"citation_id missing from crosswalk: {', '.join(missing)}")

    needed_keys = sorted({by_cid[r["citation_id"].strip()]["zotero_key"].strip() for r in active_rows})
    if any(not k for k in needed_keys):
        raise SystemExit("At least one active citation has an empty zotero_key")

    items_by_key: dict[str, dict] = {}
    api_errors: list[str] = []
    for key in needed_keys:
        try:
            items_by_key[key] = fetch_csl_json(key)
        except Exception as exc:  # noqa: BLE001
            api_errors.append(f"{key}: {exc}")
    if api_errors:
        raise SystemExit("Zotero local API validation failed:\n" + "\n".join(api_errors))
    for cid in {r['citation_id'].strip() for r in active_rows}:
        verify_identity(by_cid[cid], items_by_key[by_cid[cid]['zotero_key'].strip()])

    grouped: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in active_rows:
        grouped[row["insert_as_group"].strip()].append(row)

    def group_sort_key(group_id: str) -> tuple[int, str]:
        rows = grouped[group_id]
        order = rows[0].get("group_order", "")
        try:
            return (int(order), group_id)
        except ValueError:
            return (999999, group_id)

    payload_groups = []
    for group_id in sorted(grouped, key=group_sort_key):
        rows = grouped[group_id]
        anchors = {r.get("anchor_text", "").strip() for r in rows}
        positions = {r.get("position", "after").strip() or "after" for r in rows}
        occurrences = {r.get("anchor_occurrence", "unique").strip() or "unique" for r in rows}
        reasons = {r.get('anchor_selection_reason', '').strip() for r in rows}
        if not group_id:
            raise SystemExit('Empty insert_as_group')
        if len(anchors) != 1 or "" in anchors:
            raise SystemExit(f"Group {group_id} must have exactly one non-empty anchor_text")
        if len(positions) != 1:
            raise SystemExit(f"Group {group_id} has inconsistent position values")
        if not positions <= {'before', 'after'}:
            raise SystemExit(f'Group {group_id} has invalid position')
        if len(occurrences) != 1:
            raise SystemExit(f"Group {group_id} has inconsistent anchor_occurrence values")
        if len(reasons) != 1:
            raise SystemExit(f'Group {group_id} has inconsistent anchor_selection_reason')
        occurrence = next(iter(occurrences))
        if occurrence not in {'unique', 'first', 'last', 'after_last_in_document'} and not re.fullmatch(r'[1-9]\d*', occurrence):
            raise SystemExit(f'Group {group_id} has invalid anchor_occurrence')
        orders = [r.get('citation_order', '').strip() for r in rows]
        if any(not re.fullmatch(r'[1-9]\d*', o) for o in orders) or len(set(map(int, orders))) != len(orders):
            raise SystemExit(f'Group {group_id} requires distinct positive citation_order values')
        if len({r['citation_id'].strip() for r in rows}) != len(rows):
            raise SystemExit(f'Group {group_id} has duplicate citation_id')

        def citation_sort_key(row: dict[str, str]) -> tuple[int, str]:
            try:
                return (int(row.get("citation_order", "")), row.get("citation_id", ""))
            except ValueError:
                return (999999, row.get("citation_id", ""))

        citation_items = []
        citation_ids = []
        zotero_keys = []
        for row in sorted(rows, key=citation_sort_key):
            cid = row["citation_id"].strip()
            cw = by_cid[cid]
            key = cw["zotero_key"].strip()
            try:
                numeric_id = int(str(cw["zotero_item_id"]).strip())
            except ValueError:
                raise SystemExit(f"{cid} has invalid zotero_item_id: {cw.get('zotero_item_id')}")
            if numeric_id <= 0:
                raise SystemExit(f'{cid} has non-positive zotero_item_id')
            item_data = dict(items_by_key[key])
            uri = item_data["id"]
            item_data["id"] = numeric_id
            citation_items.append({"id": numeric_id, "uris": [uri], "itemData": item_data})
            citation_ids.append(cid)
            zotero_keys.append(key)

        citation_object = {
            "citationID": "codex-" + hashlib.sha1(group_id.encode("utf-8")).hexdigest()[:12],
            "properties": {"noteIndex": 0},
            "citationItems": citation_items,
            "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
        }
        field_code = "ZOTERO_ITEM CSL_CITATION " + json.dumps(
            citation_object, ensure_ascii=False, separators=(",", ":")
        )
        payload_groups.append(
            {
                "group_id": group_id,
                "anchor_text": next(iter(anchors)),
                "anchor_occurrence": next(iter(occurrences)),
                "anchor_selection_reason": next(iter(reasons)),
                "action": "add",
                "position": next(iter(positions)),
                "citation_ids": citation_ids,
                "zotero_keys": zotero_keys,
                "field_code": field_code,
                "placeholder": f"[{group_id}]",
            }
        )

    doc_prefs = (
        f'<data data-version="3" zotero-version="9.0.5">'
        f'<session id="{uuid.uuid4().hex[:8]}"/>'
        f'<style id={quoteattr(args.style)} hasBibliography="1" bibliographyStyleHasBeenSet="1"/>'
        f'<prefs><pref name="fieldType" value="Field"/></prefs>'
        f'</data>'
    )
    payload = {
        "doc_path": str(docx),
        "doc_sha256": hashlib.sha256(docx.read_bytes()).hexdigest(),
        "style": args.style,
        "doc_prefs": doc_prefs,
        "groups": payload_groups,
        "bibliography": {
            "heading": args.bibliography_heading,
            "field_code": 'ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY',
            "placeholder": "[Bibliography placeholder]",
        },
    }
    if args.bibliography_mode == 'omit':
        payload['bibliography'] = None
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"payload={out}")
    print(f"groups={len(payload_groups)}")
    print(f"citations={sum(len(g['citation_ids']) for g in payload_groups)}")
    print(f"unique_zotero_keys={len(needed_keys)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
