# Word Revision Highlighter

Codex skill and helper script for Microsoft Word revision manuscripts that need revised passages visibly marked with highlighter.

## 日本語での使い方

このスキルは、論文や投稿原稿の改訂版で「変更箇所を蛍光ペンで示してください」と求められたときに、CodexやClaudeなどのAIエージェントへ作業を任せるためのものです。

ユーザーがWordの比較機能やVBAマクロを細かく覚える必要はありません。旧いファイルと新しいファイル、または変更履歴つきのWordファイルを指定して、AIエージェントに次のように頼んでください。エージェントが状況を判断し、変更履歴つき文書の作成、ハイライト用マクロの生成、Wordでの実行手順の提示まで代わりに進めます。

背景の考え方は、こちらの記事も参照してください: [修正箇所を蛍光ペンで一括マークする](https://plaza.umin.ac.jp/shoei05/index.php/2023/04/23/2342/)

### こう指示してみてください

旧いファイルと新しいファイルから、変更履歴を作ってハイライトしたい場合:

```text
$word-revision-highlighter を使ってください。
旧い原稿は old_manuscript.docx、新しい原稿は revised_manuscript.docx です。
この2つを比較して変更履歴つき文書を作り、その変更履歴のうち挿入箇所を黄色でハイライトした提出用Wordファイルを作りたいです。
必要な手順とVBAマクロを作ってください。
```

すでに変更履歴があるWordファイルから、ハイライト版を作りたい場合:

```text
$word-revision-highlighter を使ってください。
manuscript_with_track_changes.docx にはすでに変更履歴があります。
挿入された本文を黄色でハイライトし、変更履歴は承認して、ハイライトだけが残る提出用のWordファイルを作りたいです。
必要なVBAマクロとWordでの実行手順を作ってください。
```

既存のハイライトを残したまま、今回の変更箇所だけ追加でハイライトしたい場合:

```text
$word-revision-highlighter を使ってください。
変更履歴つきWordファイルがあります。
既存のハイライトは消さずに、今回挿入された箇所だけを明るい緑で追加ハイライトしたいです。
VBAマクロと実行手順を作ってください。
```

削除箇所も見える形で残したい場合:

```text
$word-revision-highlighter を使ってください。
変更履歴つきWordファイルから、挿入箇所は黄色ハイライト、削除箇所は本文に戻して赤い取り消し線で表示する提出用ファイルを作りたいです。
VBAマクロと実行手順を作ってください。
```

Claudeなど英語指示のほうが通しやすいエージェントには、次のように頼めます。

```text
Use the word-revision-highlighter workflow.
I have an old Word file named old_manuscript.docx and a revised Word file named revised_manuscript.docx.
Create instructions for comparing them into a Track Changes document, then generate a VBA macro that highlights inserted revisions in yellow and accepts the revisions in a separate output document.
Please explain the steps in Japanese.
```

```text
Use the word-revision-highlighter workflow.
I already have a Word file with Track Changes.
Generate a VBA macro that preserves existing highlights, marks inserted revisions with wdBrightGreen, and creates a separate output document.
Please include Japanese instructions for running the macro in Word.
```

### バックエンドでは何をしているか

AIエージェントは、だいたい次の流れで作業します。

1. 旧いファイルと新しいファイルだけがある場合は、Wordの Review > Compare > Compare Documents を使って、旧版と新版の差分から変更履歴つき文書を作るよう案内します。
2. 変更履歴つき文書がある場合は、`scripts/generate_macro.py` で用途に合ったVBAマクロを生成します。
3. 生成されたマクロは、元ファイルを直接編集せず、アクティブなWord文書をもとに新しい文書を作ります。
4. 新しい文書内の変更履歴を順番に見て、挿入された範囲だけを `wdYellow` などの指定色でハイライトします。
5. 標準設定では、ハイライト後に変更履歴を承認し、提出しやすい見た目の文書にします。
6. オプションで、既存ハイライトの保持、変更履歴を残したままの出力、削除箇所の赤い取り消し線表示にも対応します。

内部で使うマクロ生成コマンドは次のような形です。

```bash
python3 scripts/generate_macro.py
python3 scripts/generate_macro.py --preserve-existing-highlights
python3 scripts/generate_macro.py --keep-revisions
python3 scripts/generate_macro.py --mark-deletions
python3 scripts/generate_macro.py --color wdBrightGreen
```

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
