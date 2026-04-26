#!/usr/bin/env python3
"""Generate a Word VBA macro for highlighting Track Changes insertions."""

from __future__ import annotations

import argparse
import sys


VALID_COLORS = {
    "wdYellow",
    "wdBrightGreen",
    "wdTurquoise",
    "wdPink",
    "wdGray25",
    "wdGray50",
    "wdNoHighlight",
}


def build_macro(
    *,
    color: str,
    preserve_existing_highlights: bool,
    keep_revisions: bool,
    mark_deletions: bool,
) -> str:
    if mark_deletions and keep_revisions:
        raise ValueError("--mark-deletions cannot be combined with --keep-revisions")

    lines = [
        "Sub HighlightInsertedRevisions()",
        "  Dim change As Revision",
        "  Dim outputDoc As Document",
        "",
        "  ActiveDocument.TrackRevisions = False",
        "",
        '  If ActiveDocument.Path = "" Then',
        '    MsgBox "Save the current document before running this macro."',
        "    Exit Sub",
        "  End If",
        "",
        "  If ActiveDocument.Saved = False Then",
        '    If MsgBox("Save the current document before running?", vbYesNo, "Before running") = vbYes Then',
        "      ActiveDocument.Save",
        "    Else",
        "      Exit Sub",
        "    End If",
        "  End If",
        "",
        "  Set outputDoc = Documents.Add(Template:=ActiveDocument.FullName)",
        "",
    ]

    if not preserve_existing_highlights:
        lines.extend(
            [
                "  outputDoc.Range.HighlightColorIndex = wdNoHighlight",
                "",
            ]
        )

    lines.extend(
        [
            "  For Each change In outputDoc.Revisions",
            "    Select Case change.Type",
            "      Case wdRevisionInsert",
            "        With change.Range",
            f"          .HighlightColorIndex = {color}",
        ]
    )
    if not keep_revisions:
        lines.append("          .Revisions.AcceptAll")
    lines.extend(
        [
            "        End With",
        ]
    )

    if mark_deletions:
        lines.extend(
            [
                "",
                "      Case wdRevisionDelete",
                "        With change.Range",
                "          .Revisions.RejectAll",
                "          .Font.StrikeThrough = True",
                "          .Font.Color = vbRed",
                "        End With",
            ]
        )

    if not keep_revisions:
        lines.extend(
            [
                "",
                "      Case Else",
                "        With change.Range",
                "          .Revisions.AcceptAll",
                "        End With",
            ]
        )

    lines.extend(
        [
            "    End Select",
            "  Next change",
            "",
            '  MsgBox "Finished. Review and save the new highlighted document."',
            "End Sub",
        ]
    )

    return "\n".join(lines) + "\n"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Print a Word VBA macro that highlights inserted Track Changes text."
    )
    parser.add_argument(
        "--color",
        default="wdYellow",
        help="Word WdColorIndex constant for inserted text. Default: wdYellow.",
    )
    parser.add_argument(
        "--preserve-existing-highlights",
        action="store_true",
        help="Do not clear existing highlights before marking insertions.",
    )
    parser.add_argument(
        "--keep-revisions",
        action="store_true",
        help="Highlight changed text but leave Track Changes unresolved.",
    )
    parser.add_argument(
        "--mark-deletions",
        action="store_true",
        help="Restore deleted text and format it as red strikethrough.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.color not in VALID_COLORS:
        print(
            f"error: unsupported color {args.color!r}; choose one of {', '.join(sorted(VALID_COLORS))}",
            file=sys.stderr,
        )
        return 2

    try:
        macro = build_macro(
            color=args.color,
            preserve_existing_highlights=args.preserve_existing_highlights,
            keep_revisions=args.keep_revisions,
            mark_deletions=args.mark_deletions,
        )
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    print(macro)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
