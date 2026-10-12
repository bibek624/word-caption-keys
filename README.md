# Word Caption Keys

Keyboard shortcuts for captions and cross-references in Microsoft Word.

Inserting a cross-reference in Word normally means opening a dialog, choosing the reference type, changing the format, and scrolling to the right item, every single time. With Word Caption Keys you press one key, type a few letters, and press Enter.

## Shortcuts

These only act while a Word window is active.

| Key | What it does |
|---|---|
| `Alt+F` | Lists the figures in the document. Enter inserts a cross-reference to the one you pick. |
| `Alt+T` | Lists the tables in the document. Enter inserts a cross-reference to the one you pick. |
| `Ctrl+Alt+F` | Inserts a new figure caption at the cursor. |
| `Ctrl+Alt+T` | Inserts a new table caption. |

In the list, type part of the caption or just its number to filter, move with the Up and Down arrows, and press Esc to cancel.

Cross-references are inserted as label and number only (for example "Figure 2"), as real Word fields that update when your numbering changes. Captions are real Word captions, so they renumber automatically and work with a list of figures or tables.

## Install

1. Install [AutoHotkey v2](https://www.autohotkey.com/). It is free and does not need administrator rights.
2. Download [`WordCaptionKeys.ahk`](WordCaptionKeys.ahk) (open the file, then use the download button at the top right).
3. Double-click the downloaded file. A green "H" icon appears in the system tray.

To have it start every time you sign in, right-click the tray icon and choose **Start with Windows**.

To stop it, right-click the tray icon and choose **Exit**.

## Where captions go

- On an empty line, the caption is inserted on that line.
- In a paragraph that already has text, it goes on a new line below.
- With a picture selected, it goes below the picture.
- For `Ctrl+Alt+T` with the cursor inside a table, it goes above that table.

The cursor is left right after the number, so you can type your own separator and title.

## Good to know

- Windows only, with the desktop version of Word.
- It uses Word's built-in "Figure" and "Table" caption labels. Captions made with a different label, or in a non-English version of Word, will not appear in the lists.
- While it runs, these keys stop doing their usual job in Word: `Alt+F` (File tab), `Alt+T`, `Ctrl+Alt+F` (insert footnote) and `Ctrl+Alt+T` (™ symbol). You can change the keys by editing the four lines under `#HotIf WinActive("ahk_class OpusApp")` in the script.
- The script is plain text. Open it in Notepad to see exactly what it does.

## License

[MIT](LICENSE)
