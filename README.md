# TigTag v1.0

Generates the **Tags** / **Sub-tags** columns across Excel workbooks based on
configurable keyword, gender, and file-type rules.

## Files

- `tigtag.py` — run this.
- `tigtag.png` — app icon. Keep it in the same folder as `tigtag.py`.

## Setup

```bash
pip install -r requirements.txt
python tigtag.py
```

## Building a Windows installer

```bash
pip install pyinstaller
pyinstaller tigtag.spec          # -> dist\TigTag.exe (icon + tigtag.png bundled in)
iscc tigtag.iss                  # -> installer_output\TigTag_Setup_v1.0.exe
```

Requires [Inno Setup](https://jrsoftware.org/isinfo.php) installed for `iscc`
(or open `tigtag.iss` in the Inno Setup IDE and click Compile). `tigtag.ico`
must sit alongside `tigtag.py`/`tigtag.spec`/`tigtag.iss` — it's the
multi-size icon generated from `tigtag.png`, used for the exe and installer.

## Features

- **Multi-file, multi-sheet batch processing** — load several workbooks, pick which worksheets in each to process, run them all in one pass.
- **Keyword Detection panel** — scans file/consent paths, surfaces non-generic candidate keywords with checkboxes, split into separate **Tags** and **Sub-tags** tabs. Select All / Clear All per tab, custom keyword entry, and an editable "writes as" field per keyword.
- **Blacklist & Whitelist** — keyword lists for Tags and Sub-tags, each importable/exportable as JSON.
- **Male/Female auto-detection** — literal gender tokens in the path first, then a fallback custom name list (importable/exportable). Rows with no match simply get no gender tag.
- **File-type validation** — only rows with a recognized extension are tagged as files; the extension list is editable. Folder-only rows are skipped by default (toggle to include them).
- **Output** — saved next to the source file as `<name>_processed.xlsx`; the source file is never modified.
- **Feedback popup** after each batch completes.
- **About dialog** with version and copyright info.

## Output location

Each processed file is written alongside its source as `<original_name>_processed.xlsx`.

## Local data

Keyword rules, blacklist/whitelist, name lists, extensions, and feedback are stored locally in `~/.tigtag/` — nothing leaves your machine.

---
© 2026 Kiran Mukhiya. All rights reserved.
