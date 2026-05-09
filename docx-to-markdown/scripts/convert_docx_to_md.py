#!/usr/bin/env python3
"""Convert a .docx file to UTF-8 Markdown using only the Python stdlib."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{WORD_NS}}}"
NS = {"w": WORD_NS}


def paragraph_text(paragraph: ET.Element) -> str:
    parts: list[str] = []
    for node in paragraph.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag == W + "tab":
            parts.append("    ")
        elif node.tag == W + "br":
            parts.append("\n")
    return "".join(parts).replace("\xa0", " ")


def cell_text(cell: ET.Element) -> str:
    paragraphs: list[str] = []
    for paragraph in cell.findall("./w:p", NS):
        text = paragraph_text(paragraph).strip()
        if text:
            paragraphs.append(text)
    return "<br>".join(paragraphs)


def escape_cell(text: str) -> str:
    return text.replace("|", r"\|").replace("\n", "<br>")


def table_to_markdown(table: ET.Element) -> str:
    rows: list[list[str]] = []
    for table_row in table.findall("./w:tr", NS):
        row = [
            escape_cell(cell_text(cell).strip())
            for cell in table_row.findall("./w:tc", NS)
        ]
        if any(row):
            rows.append(row)

    if not rows:
        return ""

    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]

    lines = [
        "| " + " | ".join(rows[0]) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def convert_docx_to_markdown(source: Path) -> str:
    with ZipFile(source) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))

    body = root.find("w:body", NS)
    if body is None:
        return ""

    output: list[str] = []
    first_heading_done = False
    previous_blank = True

    for child in list(body):
        if child.tag == W + "p":
            text = paragraph_text(child).strip()
            if not text:
                previous_blank = True
                continue

            if not first_heading_done:
                output.append("# " + text)
                first_heading_done = True
            else:
                if not previous_blank:
                    output.append("")
                output.append(text)
            previous_blank = False

        elif child.tag == W + "tbl":
            markdown = table_to_markdown(child)
            if markdown:
                if output and output[-1] != "":
                    output.append("")
                output.append(markdown)
                previous_blank = False

    return "\n".join(output).rstrip() + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert a .docx file to UTF-8 Markdown."
    )
    parser.add_argument("source", type=Path, help="Path to the .docx file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .md path. Defaults to source path with .md extension.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source = args.source
    output = args.output or source.with_suffix(".md")

    markdown = convert_docx_to_markdown(source)
    output.write_text(markdown, encoding="utf-8")

    print(output)
    print(f"chars {len(markdown)} lines {markdown.count(chr(10))}")


if __name__ == "__main__":
    main()
