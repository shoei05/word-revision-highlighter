# Word Revision Highlighter

Codex skill and helper script for Microsoft Word revision manuscripts that need revised passages visibly marked with highlighter.

## 日本語での使い方

このスキルは、論文や投稿原稿の改訂版で「変更箇所を蛍光ペンで示してください」と求められたときに使います。

主に2つの場面を想定しています。

1. 旧いWordファイルと新しいWordファイルから、変更履歴つき文書を作る。
2. すでに変更履歴があるWordファイルから、挿入された本文をハイライトした提出用文書を作る。

### 1. 旧いファイルと新しいファイルから履歴を作る

変更履歴を記録していなかった場合でも、Wordの比較機能を使うと、旧版と新版の差分から変更履歴つき文書を作れます。

手順:

1. Wordで旧いファイルまたは新しいファイルを開きます。
2. Review > Compare > Compare Documents を選びます。
3. Original document に旧いファイルを指定します。
4. Revised document に新しいファイルを指定します。
5. 比較結果として作られた文書を保存します。
6. その比較結果の文書に対して、このスキルのハイライト用マクロを実行します。

### 2. 履歴がある文書からハイライトを入れる

すでにTrack Changes、つまり変更履歴が入っているWordファイルがある場合は、その文書を開いた状態でマクロを実行します。

標準のマクロは次の処理をします。

- 元ファイルを直接壊さないように、新しい文書を作る。
- 既存のハイライトを消す。
- 挿入された文字列を黄色でハイライトする。
- 変更履歴を承認して、ハイライトだけが残る提出用文書にする。

既存のハイライトを残したい場合、変更履歴も表示したままにしたい場合、削除箇所を赤い取り消し線で残したい場合は、下のオプションを使います。

```bash
python3 scripts/generate_macro.py --preserve-existing-highlights
python3 scripts/generate_macro.py --keep-revisions
python3 scripts/generate_macro.py --mark-deletions
```

### Codexへの指示例

旧版と新版から比較文書を作りたい場合:

```text
$word-revision-highlighter を使ってください。
旧いファイルは old_manuscript.docx、新しいファイルは revised_manuscript.docx です。
この2つからWordの比較機能で変更履歴つき文書を作り、その後、挿入箇所を黄色でハイライトする提出用ファイルを作る手順とVBAマクロをください。
```

すでに変更履歴つき文書がある場合:

```text
$word-revision-highlighter を使ってください。
manuscript_with_track_changes.docx にはすでに変更履歴があります。
挿入された本文を黄色でハイライトし、変更履歴は承認して、提出用のきれいな文書にしたいです。必要なVBAマクロとWordでの実行手順をください。
```

既存のハイライトを残したい場合:

```text
$word-revision-highlighter を使ってください。
変更履歴つきのWordファイルがあります。既存のハイライトは消さずに、今回の挿入箇所だけを明るい緑で追加ハイライトするVBAマクロを作ってください。
```

削除箇所も見えるようにしたい場合:

```text
$word-revision-highlighter を使ってください。
変更履歴つき文書から、挿入箇所は黄色ハイライト、削除箇所は本文に戻して赤い取り消し線で表示する提出用ファイルを作りたいです。VBAマクロと実行手順をください。
```

### Claudeへの指示例

```text
Use the word-revision-highlighter workflow.
I have an old Word file named old_manuscript.docx and a revised Word file named revised_manuscript.docx.
First, explain how to create a Track Changes document using Word's Compare Documents feature.
Then generate a VBA macro that highlights inserted revisions in yellow and accepts the revisions in a new output document.
Please provide the steps in Japanese.
```

```text
Use the word-revision-highlighter workflow.
I already have a Word file with Track Changes.
Generate a VBA macro that preserves existing highlights, marks inserted revisions with wdBrightGreen, and creates a separate output document.
Please include Japanese instructions for running the macro in Word.
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
