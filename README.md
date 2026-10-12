# Word Caption Keys

Keyboard shortcuts for captions and cross-references in Microsoft Word.

Inserting a cross-reference in Word normally means opening a dialog, choosing the reference type, changing the format, and scrolling to the right item, every single time. With Word Caption Keys you press one key, type a few letters, and press Enter.

## Shortcuts

These only act while a Word window is active.

| Key | What it does |
|---|---|
| `Alt+F` | Lists the figures in the document. Enter inserts a cross-reference to the one you pick. |
| `Alt+T` | Lists the tables in the document. Enter inserts a cross-reference to the one you pick. |
| `Alt+H` | Lists the headings in the document. Enter inserts a cross-reference to the one you pick. |
| `Ctrl+Alt+F` | Inserts a new figure caption at the cursor. |
| `Ctrl+Alt+T` | Inserts a new table caption. |
| `Ctrl+Alt+U` | Updates every field in the document. |

In a list, type part of the text or just the number to filter, move with the Up and Down arrows, press Enter to insert, and press Esc to cancel.

## Install

You do this once. It takes about two minutes and does not need administrator rights.

### 1. Install AutoHotkey

Word Caption Keys runs on a free program called AutoHotkey.

1. Go to [autohotkey.com](https://www.autohotkey.com/) and click **Download**, then **Download v2.0**.
2. Open the downloaded file and click **Install**.

### 2. Download Word Caption Keys

1. Open [`WordCaptionKeys.ahk`](WordCaptionKeys.ahk) on this page.
2. Click the download button (the arrow pointing down, at the top right of the file).
3. Move the downloaded file to a folder where it can stay, such as your Documents folder. Do this before running it: if you move the file later, automatic startup stops working until you switch it on again.

### 3. Run it

Double-click `WordCaptionKeys.ahk`.

A welcome message lists the shortcuts and asks whether to start Word Caption Keys automatically every time Windows starts. Click **Yes** and you never have to think about it again.

That's it. Open a Word document and press `Alt+F`.

### If double-clicking opens Notepad, or Windows asks which app to use

This means Windows has not yet linked `.ahk` files to AutoHotkey.

1. Right-click `WordCaptionKeys.ahk` and choose **Open with**, then **Choose another app**.
2. Pick **AutoHotkey** from the list. If it is not listed, scroll down and choose **Choose an app on your PC**, then browse to the AutoHotkey folder and pick `AutoHotkey64.exe` inside the `v2` folder.
3. Click **Always**, so that double-clicking works from now on.

### If Windows shows a security warning

Because the file came from the internet, Windows may ask whether you really want to open it. Click **Open** or **Run**. The script is plain text, so you can open it in Notepad first to see exactly what it does.

## The tray icon

While Word Caption Keys is running, a green **H** icon sits near the clock at the bottom right of the screen. If you do not see it, click the small **^** arrow there to show hidden icons.

Right-click the icon for:

- **Shortcuts**: a reminder of the keys.
- **Settings**: opens the settings file.
- **Reload**: applies changes you made to the settings.
- **Start with Windows**: switches automatic startup on or off. A tick means it is on.
- **Exit**: stops Word Caption Keys.

## How it behaves

**Cross-references to figures and tables** are inserted as label and number only, for example "Figure 2". They are real Word fields, so they update when your numbering changes, and Ctrl+click jumps to the figure or table.

**Cross-references to headings** insert the heading number, for example "3.2", so you can write "see Section 3.2". If the heading has no number, its text is inserted instead. Press `Ctrl+Enter` in the list to insert the heading text even when it has a number.

**New captions** are real Word captions, so they renumber automatically and work with a list of figures or tables. The cursor is left right after the number, so you can type your own separator and title. Where the caption goes:

- On an empty line, it is inserted on that line.
- In a paragraph that already has text, it goes on a new line below.
- With a picture selected, it goes below the picture.
- For a table caption with the cursor inside a table, it goes above that table.

**Update all fields** refreshes cross-references, caption numbers, tables of contents and page numbers throughout the document, including headers, footers and footnotes. Word's own `F9` key only updates what you have selected.

## Changing the keys

The default keys replace a few of Word's own shortcuts while Word Caption Keys is running: `Alt+F` (File tab), `Alt+H` (Home tab key tips), `Ctrl+Alt+F` (insert footnote) and `Ctrl+Alt+T` (™ symbol). If you use any of those, change the keys.

1. Right-click the tray icon and choose **Settings**. This opens `WordCaptionKeys.ini`, which sits next to the script.
2. Change the key after the `=` sign. The symbols are: `!` for Alt, `^` for Ctrl, `+` for Shift, `#` for the Windows key. For example, `^+f` means Ctrl+Shift+F.
3. Save the file, then right-click the tray icon and choose **Reload**.

To switch a shortcut off, leave nothing after its `=` sign.

```ini
[Keys]
FigureReference=!f
TableReference=!t
HeadingReference=!h
FigureCaption=^!f
TableCaption=^!t
UpdateAllFields=^!u
```

## Caption labels

Word's built-in Figure and Table labels are found automatically, whatever they are called in your language version of Word.

If you created your own caption labels, any whose names start with "Fig" or "Tab" (for example "Fig." or "Tab.") are included in the figure and table lists too. To include labels with other names, add them to the `[Labels]` section of the settings file, separated by commas:

```ini
[Labels]
FigureLabels=Fig,Chart,Photo
TableLabels=Tab
```

When you insert a new caption, the label your document already uses most is the one that gets used, so the numbering stays in one sequence.

## Good to know

- Windows only, with the desktop version of Word.
- Figures and tables appear in the lists only if they have real Word captions (made with these shortcuts or with **References > Insert Caption**). Numbers typed by hand are not found.
- Headings appear only if they use Word's heading styles.
- If Track Changes is on, inserted references and updated fields are recorded as changes, as they would be if you did it by hand.

## Uninstall

Right-click the tray icon, untick **Start with Windows**, and choose **Exit**. Then delete `WordCaptionKeys.ahk` and `WordCaptionKeys.ini`.

## License

[MIT](LICENSE)
