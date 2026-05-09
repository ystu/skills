---
name: docx-to-markdown
description: Convert Microsoft Word .docx files to UTF-8 Markdown while preserving readable paragraphs and Markdown tables. Use when Codex is asked to turn a .docx into .md, extract Word document contents into Markdown, or handle Chinese/Unicode Word documents without relying on pandoc or python-docx.
---

# DOCX to Markdown

Use this skill when converting `.docx` files into Markdown, especially in workspaces with Chinese filenames or Chinese document content.

## Workflow

1. Locate the target `.docx`.
   - If the user provides a path, use that path.
   - If PowerShell or Python mangles a Chinese path, locate the file with `Get-ChildItem` or a glob such as `Path("doc").glob("*.docx")`.

2. Prefer the bundled converter:

```powershell
python "C:\Users\ADMIN\.codex\skills\docx-to-markdown\scripts\convert_docx_to_md.py" "path\to\file.docx"
```

3. If the skill is being used from a copied or local skill folder, run the script from that folder instead:

```powershell
python "path\to\docx-to-markdown\scripts\convert_docx_to_md.py" "path\to\file.docx"
```

4. Verify the result:
   - Read the first 40-80 lines with explicit UTF-8 in PowerShell:

```powershell
Get-Content -Encoding UTF8 -LiteralPath "path\to\file.md" -TotalCount 80
```

   - Count Markdown tables if the source document had tables:

```powershell
Select-String -Encoding UTF8 -LiteralPath "path\to\file.md" -Pattern '^\| ---' | Measure-Object
```

## Converter Behavior

The script reads `.docx` as a ZIP archive and parses `word/document.xml` directly with Python standard-library modules. It does not require `pandoc`, `python-docx`, network access, or package installation.

It preserves:

- UTF-8 text output.
- Paragraph order from the Word body.
- Word tables as Markdown pipe tables.
- In-cell line breaks as `<br>`.
- Tabs and soft line breaks where practical.

It intentionally does not preserve visual styling, images, headers/footers, tracked changes, comments, or complex merged-cell layout. For poster or planning workflows, treat the Markdown as an editable text extraction, not a pixel-perfect Word clone.

## Output Naming

By default, write the Markdown next to the source file with the same base name and `.md` extension. Use `--output` when the user requests a different filename.

Examples:

```powershell
python "C:\Users\ADMIN\.codex\skills\docx-to-markdown\scripts\convert_docx_to_md.py" "doc\活動規劃.docx"
python "C:\Users\ADMIN\.codex\skills\docx-to-markdown\scripts\convert_docx_to_md.py" "doc\活動規劃.docx" --output "doc\活動規劃.md"
```
