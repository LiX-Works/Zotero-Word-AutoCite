#!/usr/bin/env python3
"""Validate a Zotero import manifest and create a crosswalk seed.

This script does not modify Zotero. It is an audit/planning step before a
project-specific Zotero import script is generated or run.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path


REQUIRED = {"citation_id", "title", "doi", "year", "authors", "journal", "pdf_path"}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def norm_doi(value: str) -> str:
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", (value or "").strip(), flags=re.I).lower()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--crosswalk-seed", required=True)
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    manifest = Path(args.manifest)
    rows = read_csv(manifest)
    if not rows:
        raise SystemExit(f"manifest is empty: {manifest}")
    missing_cols = sorted(REQUIRED - set(rows[0].keys()))
    if missing_cols:
        raise SystemExit("manifest missing required columns: " + ", ".join(missing_cols))

    seen_cid: dict[str, int] = {}
    seen_doi: dict[str, int] = {}
    issues = []
    seed_rows = []
    for idx, row in enumerate(rows, start=2):
        cid = (row.get("citation_id") or "").strip()
        doi = norm_doi(row.get("doi") or "")
        pdf = Path((row.get("pdf_path") or "").strip())
        row_issues = []
        if not cid:
            row_issues.append("missing citation_id")
        if cid and cid in seen_cid:
            row_issues.append(f"duplicate citation_id also on row {seen_cid[cid]}")
        if cid:
            seen_cid[cid] = idx
        if not doi:
            row_issues.append("missing DOI")
        if doi and doi in seen_doi:
            row_issues.append(f"duplicate DOI also on row {seen_doi[doi]}")
        if doi:
            seen_doi[doi] = idx
        for field in ["title", "year", "authors", "journal"]:
            if not (row.get(field) or "").strip():
                row_issues.append(f"missing {field}")
        if not pdf.is_file():
            row_issues.append(f"pdf_path does not exist: {pdf}")
        elif pdf.suffix.lower() != ".pdf":
            row_issues.append(f"pdf_path is not a PDF: {pdf}")
        if row_issues:
            issues.append({"row": idx, "citation_id": cid, "doi": doi, "issues": row_issues})
        seed_rows.append(
            {
                "citation_id": cid,
                "aliases": "",
                "title": (row.get("title") or "").strip(),
                "doi": doi,
                "year": (row.get("year") or "").strip(),
                "first_author": ((row.get("authors") or "").split(";")[0]).strip(),
                "journal": (row.get("journal") or "").strip(),
                "pdf_path": str(pdf),
                "zotero_key": "",
                "zotero_item_id": "",
                "zotero_collection": (row.get("collection") or "").strip(),
                "source_metadata_file": str(manifest),
                "notes": "pending_zotero_import",
            }
        )

    out = Path(args.crosswalk_seed)
    out.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "citation_id",
        "aliases",
        "title",
        "doi",
        "year",
        "first_author",
        "journal",
        "pdf_path",
        "zotero_key",
        "zotero_item_id",
        "zotero_collection",
        "source_metadata_file",
        "notes",
    ]
    with out.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(seed_rows)

    report = {
        "manifest": str(manifest),
        "row_count": len(rows),
        "crosswalk_seed": str(out),
        "issue_count": len(issues),
        "issues": issues,
        "ready_for_import": len(issues) == 0,
        "note": "This script does not modify Zotero. Continue within existing import authorization when checks pass; ask about unresolved material conflicts or missing authority.",
    }
    report_path = Path(args.report)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if len(issues) == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
