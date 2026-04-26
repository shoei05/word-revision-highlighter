---
name: word-revision-highlighter
description: Create and customize Microsoft Word VBA workflows for academic manuscript revision files that need changed text marked with highlighter. Use when Codex needs to help with Word Track Changes, revision manuscripts, highlighted changes for journal resubmission, "変更箇所を蛍光ペン", "修正履歴をハイライト", comparing old/new Word files to reconstruct revisions, preserving existing highlights, keeping or accepting revision marks, changing highlight colors, or marking deletions as red strikethrough.
---

# Word Revision Highlighter

## Core Workflow

Use this skill to help users produce a Word manuscript file where revised passages are visibly marked for journal or editor resubmission.

1. Confirm the available source:
   - If the user has a Word file with Track Changes, use it directly.
   - If the user has only old and revised Word files, instruct Word's Review > Compare > Compare Documents workflow first to create a document with revisions.
2. Decide the desired output:
   - Default: create a new document, remove previous highlights, highlight inserted text, and accept revision marks.
   - Preserve existing highlights: omit the global highlight-clearing step.
   - Keep Track Changes visible: do not accept/reject revisions after marking.
   - Show deletions: reject deleted ranges back into text, then format them red with strikethrough.
   - Use another color: change the `WdColorIndex` constant, such as `wdBrightGreen`, `wdPink`, or `wdTurquoise`.
3. Generate or adapt the macro. Prefer `scripts/generate_macro.py` for a clean starting point.
4. Give the user Word execution steps:
   - Save the Word file first.
   - Open Tools > Macro > Visual Basic Editor, or Developer > Visual Basic.
   - Insert > Module.
   - Paste the macro.
   - Run the macro from the active revision document.
   - Save the generated document under a new filename before submission.

## Script

Run the generator when the user wants the VBA code or a variant:

```bash
python3 scripts/generate_macro.py --help
```

Common examples:

```bash
python3 scripts/generate_macro.py
python3 scripts/generate_macro.py --preserve-existing-highlights --color wdBrightGreen
python3 scripts/generate_macro.py --keep-revisions
python3 scripts/generate_macro.py --mark-deletions
```

## Reference

Read `references/workflow-notes.md` when the user asks for:

- Explanation of what each macro section does.
- Guidance for reconstructing Track Changes from old/new files.
- Color constants.
- Tradeoffs between accepting revisions, keeping revisions, and marking deletions.

Source inspiration: https://plaza.umin.ac.jp/shoei05/index.php/2023/04/23/2342/
