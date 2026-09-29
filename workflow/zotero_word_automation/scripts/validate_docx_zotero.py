#!/usr/bin/env python3
"""Validate Zotero fields in a docx after insertion/refresh."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


def read_docx_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def xml_to_text(xml: str) -> str:
    return inspect_xml(xml)[0]


def inspect_xml(xml: str) -> tuple[str, list[dict], list[str]]:
    """Keep field instructions separate from their displayed results."""
    w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    stack, fields, visible, errors = [], [], [], []

    def emit(text):
        if all(f['phase'] == 'result' for f in stack):
            visible.append(text)
            for f in stack:
                f['result'] += text

    def walk(node):
        if node.tag in {w + 'del', w + 'moveFrom'}:
            return
        if node.tag == w + 'r':
            for name in ('vanish', 'webHidden'):
                prop = node.find(f'{w}rPr/{w}{name}')
                if prop is not None and prop.get(w + 'val', 'true') not in {'0', 'false', 'off'}:
                    return
        if node.tag == w + 'fldSimple':
            stack.append({'code': node.get(w + 'instr', ''), 'result': '', 'phase': 'result'})
            for child in node:
                walk(child)
            fields.append(stack.pop())
            return
        if node.tag == w + 'fldChar':
            kind = node.get(w + 'fldCharType')
            if kind == 'begin':
                stack.append({'code': '', 'result': '', 'phase': 'code'})
            elif kind == 'separate' and stack:
                stack[-1]['phase'] = 'result'
            elif kind == 'end' and stack:
                fields.append(stack.pop())
            else:
                errors.append('invalid field delimiter')
        elif node.tag == w + 'instrText' and stack:
            stack[-1]['code'] += node.text or ''
        elif node.tag == w + 't':
            emit(node.text or '')
        elif node.tag in {w + 'tab', w + 'br'}:
            emit(' ')
        for child in node:
            walk(child)
        if node.tag == w + 'p':
            emit('\n')

    walk(ET.fromstring(xml))
    if stack:
        errors.append('unclosed fields')
    return ' '.join(''.join(visible).split()), fields, errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True)
    parser.add_argument("--payload", help="Optional payload JSON for expected group count")
    parser.add_argument("--payload-scope", choices=["full", "subset"], default="full")
    parser.add_argument("--expected-items", type=int)
    parser.add_argument("--expected-bibs", type=int)
    parser.add_argument("--bibliography-contains", action="append", default=[])
    parser.add_argument("--bibliography-not-contains", action="append", default=[])
    parser.add_argument("--output", help="Optional JSON output path")
    parser.add_argument("--must-contain", action="append", default=[], help="Text that must appear")
    parser.add_argument("--must-not-contain", action="append", default=[], help="Text that must not appear")
    args = parser.parse_args()

    docx = Path(args.docx)
    if not docx.exists():
        raise SystemExit(f"docx does not exist: {docx}")
    xml = read_docx_xml(docx)
    text, fields, errors = inspect_xml(xml)
    items = [f for f in fields if 'ZOTERO_ITEM CSL_CITATION' in f['code']]
    bibs = [f for f in fields if 'ZOTERO_BIBL' in f['code']]
    bib_text = ' '.join(' '.join(f['result'] for f in bibs).split())
    parsed = {}
    for f in items:
        try:
            obj = json.loads(f['code'][f['code'].index('{'):].strip())
            cid = obj['citationID']
            if cid in parsed:
                errors.append(f'duplicate citationID: {cid}')
            parsed[cid] = obj
        except (ValueError, KeyError, TypeError) as exc:
            errors.append(f'invalid citation JSON: {exc}')
    if any(not f['result'].strip() for f in items + bibs):
        errors.append('empty citation or bibliography display')
    if len(bibs) > 1:
        errors.append('multiple bibliography fields; check whether intended')

    result = {
        "docx": str(docx),
        "zotero_item_field_count": len(items),
        "zotero_bibliography_field_count": len(bibs),
        "placeholder_like_text_found": bool(re.search(r"\[(?:[^\]\r\n]*placeholder[^\]\r\n]*|(?:TG\d+|CL\d+)[\w.\-]*|[\w.\-]+_G\d+)\]", text, re.I)),
        "visible_bibliography_text": bib_text,
        "refresh_completion": "not_determined_by_static_check",
        "limitations": "Main document XML only; layout, inherited hidden styles, notes/headers and live Zotero completion require separate checks.",
        "must_contain": {},
        "must_not_contain": {},
        "warnings": errors,
    }

    if args.payload:
        payload = json.loads(Path(args.payload).read_text(encoding="utf-8"))
        result["payload_group_count"] = len(payload.get("groups", []))
        result["payload_citation_count"] = sum(len(g.get("citation_ids", [])) for g in payload.get("groups", []))
        if args.payload_scope == 'full' and result["zotero_item_field_count"] != result["payload_group_count"]:
            result["warnings"].append("ZOTERO_ITEM field count differs from payload group count")
        for group in payload.get('groups', []):
            expected = json.loads(group['field_code'][group['field_code'].index('{'):])
            actual = parsed.get(expected['citationID'])
            if not actual or [i.get('uris') for i in actual['citationItems']] != [i.get('uris') for i in expected['citationItems']]:
                errors.append(f'missing or changed citation members/order: {group["group_id"]}')
        placeholders = [g.get('placeholder', '') for g in payload.get('groups', [])]
        placeholders.append((payload.get('bibliography') or {}).get('placeholder', ''))
        if any(p and p in text for p in placeholders):
            result['placeholder_like_text_found'] = True

    if result['placeholder_like_text_found']:
        errors.append('placeholder-like visible text remains')
    for name, actual, expected in [('items', len(items), args.expected_items), ('bibliographies', len(bibs), args.expected_bibs)]:
        if expected is not None and actual != expected:
            errors.append(f'expected {expected} {name}, found {actual}')
    for needle in args.bibliography_contains:
        if ' '.join(needle.split()) not in bib_text:
            errors.append(f'missing visible bibliography text: {needle}')
    for needle in args.bibliography_not_contains:
        if ' '.join(needle.split()) in bib_text:
            errors.append(f'forbidden visible bibliography text: {needle}')

    for needle in args.must_contain:
        result["must_contain"][needle] = needle in text
        if needle not in text:
            result["warnings"].append(f"missing required text: {needle}")

    for needle in args.must_not_contain:
        result["must_not_contain"][needle] = needle not in text
        if needle in text:
            result["warnings"].append(f"forbidden text found: {needle}")

    result['static_checks_passed'] = not result['warnings']
    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not result["warnings"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
