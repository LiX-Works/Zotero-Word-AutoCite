#!/usr/bin/env python3
"""Create a conservative inventory of PDF folders for Zotero import planning.

This does not identify papers from PDFs. It only records files and lightweight
filename guesses so Codex/user can build a verified metadata manifest.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import re
from pathlib import Path


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def guess_from_name(path: Path) -> dict[str, str]:
    stem = path.stem
    parts = stem.split("_")
    guess = {
        "guessed_year": "",
        "guessed_first_author": "",
        "guessed_title": "",
        "guessed_journal": "",
        "guessed_doi_suffix": "",
    }
    if parts and re.fullmatch(r"\d{4}", parts[0]):
        guess["guessed_year"] = parts[0]
    if len(parts) > 1:
        guess["guessed_first_author"] = parts[1]
    if len(parts) > 2:
        guess["guessed_title"] = parts[2]
    if len(parts) > 3:
        guess["guessed_journal"] = parts[3]
    if len(parts) > 4:
        guess["guessed_doi_suffix"] = "_".join(parts[4:])
    return guess


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf-folder", action="append", required=True, help="Folder to scan; repeatable")
    parser.add_argument("--output", required=True, help="Output CSV")
    parser.add_argument("--hash", action="store_true", help="Compute SHA256; slower")
    args = parser.parse_args()

    rows = []
    for folder_text in args.pdf_folder:
        folder = Path(folder_text)
        if not folder.exists():
            raise SystemExit(f"PDF folder does not exist: {folder}")
        for pdf in sorted(folder.rglob("*.pdf")):
            stat = pdf.stat()
            row = {
                "pdf_path": str(pdf),
                "folder": str(folder),
                "file_name": pdf.name,
                "size_bytes": stat.st_size,
                "sha256": sha256_file(pdf) if args.hash else "",
            }
            row.update(guess_from_name(pdf))
            rows.append(row)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    fields = [
        "pdf_path",
        "folder",
        "file_name",
        "size_bytes",
        "sha256",
        "guessed_year",
        "guessed_first_author",
        "guessed_title",
        "guessed_journal",
        "guessed_doi_suffix",
    ]
    with out.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"output={out}")
    print(f"pdf_count={len(rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
