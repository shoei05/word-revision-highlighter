# Word Revision Highlighter

Codex skill and helper script for Microsoft Word revision manuscripts that need revised passages visibly marked with highlighter.

The main workflow generates a VBA macro that:

- creates a new document from the active Word file;
- highlights inserted Track Changes text;
- optionally preserves existing highlights;
- optionally keeps revision marks visible;
- optionally restores deleted passages as red strikethrough text.

## Usage

Generate the default macro:

```bash
python3 scripts/generate_macro.py
```

Show available options:

```bash
python3 scripts/generate_macro.py --help
```

Examples:

```bash
python3 scripts/generate_macro.py --preserve-existing-highlights --color wdBrightGreen
python3 scripts/generate_macro.py --keep-revisions
python3 scripts/generate_macro.py --mark-deletions
```

## Word Steps

1. Save the Word document with Track Changes.
2. Open Tools > Macro > Visual Basic Editor, or Developer > Visual Basic.
3. Insert a new module.
4. Paste the generated macro.
5. Run the macro from the active revision document.
6. Review and save the generated document under a new filename.

If you only have old and revised Word files, use Word's Review > Compare > Compare Documents first, then run the macro on the compared document.

## Files

- `SKILL.md`: Codex skill instructions.
- `scripts/generate_macro.py`: macro generator.
- `references/workflow-notes.md`: workflow notes, color constants, and cautions.
- `agents/openai.yaml`: skill display metadata.

Source inspiration: https://plaza.umin.ac.jp/shoei05/index.php/2023/04/23/2342/
