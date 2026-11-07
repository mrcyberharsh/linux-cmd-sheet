"""Search and inspect the DOCX Linux command sheet from the terminal."""
from __future__ import annotations
import argparse
from pathlib import Path
from typing import Any

from docx import Document


def _iter_entries(document):
    section = "General"
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        style = getattr(paragraph.style, "name", "")
        if style.startswith("Heading"):
            section = text
        yield {"section": section, "text": text}
    for table_index, table in enumerate(document.tables, start=1):
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if cells:
                yield {"section": f"Table {table_index}", "text": " | ".join(cells)}


def list_sections(path: str | Path) -> list[str]:
    document = Document(str(path))
    sections = []
    for paragraph in document.paragraphs:
        if getattr(paragraph.style, "name", "").startswith("Heading") and paragraph.text.strip():
            sections.append(paragraph.text.strip())
    return sections


def search_commands(path: str | Path, query: str, limit: int = 20) -> list[dict[str, str]]:
    if not query.strip():
        raise ValueError("Search query cannot be empty.")
    if limit < 1 or limit > 200:
        raise ValueError("Limit must be between 1 and 200.")
    document = Document(str(path))
    needle = query.casefold()
    matches = []
    for entry in _iter_entries(document):
        if needle in entry["text"].casefold():
            matches.append(entry)
            if len(matches) >= limit:
                break
    return matches


def main() -> int:
    parser = argparse.ArgumentParser(description="Search the MR CYBER Linux Command Sheet DOCX.")
    parser.add_argument("query", nargs="?", help="Text or command to search for.")
    parser.add_argument("--file", default="Linux_Command_Sheet_Cybersecurity_MrCyberHarsh.docx")
    parser.add_argument("--list-sections", action="store_true", help="List document headings.")
    parser.add_argument("--limit", type=int, default=20, help="Maximum matches (1–200).")
    args = parser.parse_args()
    if not Path(args.file).is_file():
        parser.error(f"DOCX file not found: {args.file}")
    if args.list_sections:
        for section in list_sections(args.file):
            print(section)
        return 0
    if not args.query:
        parser.error("Provide a search query or use --list-sections.")
    try:
        results = search_commands(args.file, args.query, args.limit)
    except ValueError as exc:
        parser.error(str(exc))
    if not results:
        print("No matches found.")
        return 0
    for item in results:
        print(f"[{item['section']}] {item['text']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
