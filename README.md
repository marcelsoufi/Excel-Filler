# Excel Report Filler

A small desktop app that fills a report workbook from source sheets **using cell mappings you define once and reuse**.
You pick the source and target cells (by clicking or typing addresses), save the mapping as a named **template**, and apply it to the next file in seconds. The original file is never modified: the result is saved as a new copy beside it.

Built with Python, [PySide6](https://doc.qt.io/qtforpython-6/) and [openpyxl](https://openpyxl.readthedocs.io/).

## Features

- **Manual, predictable mapping** – you decide exactly which cell goes where. No guessing, no fuzzy matching.
- **Click or type** – select cells directly in the built-in sheet preview, or type addresses such as `Survey!B5:B40 ; 'Sheet 2'!C3`. Mix both freely.
- **Multiple source sheets and files** – sources can come from any sheet of the workbook and, if you want, from other Excel files (*Source files* button). Targets are always in the file you are filling.
- **Sequential fill** – fill a whole table instead of typing it by hand: map the cells of ONE sample source file to the first record of the table, then add all your source files (or a folder). The same cells are read from every file and each file goes to the next row (or column), optionally numbering the records and continuing below the records already in the table. A preview shows the final table and flags sheet mismatches or partly empty records before anything is written. If the app picked the wrong sheet in a source file, double-click that file in the preview and choose the right one (optionally for all files with the same sheet name). A record is skipped only when ALL its mapped cells are empty; a partly filled record still takes its row, with a warning.
- **Pair by clicking** – switch on *Pair by clicking*, click a source cell, then the exact target cell it must go to. The pair is added instantly, so you decide every correspondence in any order (no sequence or randomness involved).
- **Strict one-to-one pairing** – every source cell is paired with exactly one target cell, in the same order (row by row or column by column, your choice). A preview table shows every pair before it is added, and mismatched counts are rejected. Several source cells can also feed ONE target cell (first non-empty value, join as lines, or sum).
- **Reusable templates** – save the ordered list of mappings under a name, then pick it when you open another file. No need to select the cells again. If sheet names differ, the app asks which sheet to use for each one (with smart suggestions) and remembers your choice. Templates also remember the location of any extra source files they use and re-attach them automatically (or ask you to locate them if they moved).
- **Several files at once** – each file gets its own independent panel. On *Execute* the app asks whether to apply the same operation to **all** open files, to **some** of them (you pick), or **only** the current one.
- **Light and dark mode** – switch with one button; your choice is remembered.
- **Live preview before writing** – see the value that will be written, the status of every mapping (ready, overwrites, source empty, sheet missing, duplicate target…), and override any value by hand for the current file.
- **Safe export** – writes `<name>_filled.xlsx` next to the original; the original is left untouched.
- Optional conversion of Arabic-Indic digits (٣٥٠) to Latin (350), merged-cell aware, `.xlsx` and `.xlsm` support.

## Installation

Requires Python 3.9+.

```bash
pip install PySide6 openpyxl
python excel_report_filler.py
```

Keep `excel_report_filler.ico` in the same folder as the script to get the app icon.

## Usage

1. **Add files** – click *Add Excel files* or drag `.xlsx` / `.xlsm` files into the window.
2. **Select cells** – in the sheet preview, select the source cells, then press **Set** (or **Add** to append cells from another sheet) next to *Source*. Do the same for *Target*. You can also type the addresses.
3. **Add mapping** – the pair is added to the list. Blue cells in the preview are sources, green cells are targets.
4. **Save as template** – give the mapping list a name. Next time, choose it from the *Template* menu when you open a file.
5. **Execute & Save new copy** – review the preview, then export.

### Using cells from other files

Click **Source files ▾** and either pick one of the other open files or browse for a file. Its sheets appear as extra tabs (`Other.xlsx › Survey`) in the preview, and you select cells there like any other sheet. In typed addresses they are written `'[Other.xlsx]Survey'!B5`. Source files are only read, never modified.

### Applying the same operation to several files

When more than one file is open, *Execute* asks: **Yes, all open files** · **Some files…** (you tick the files) · **No, only this file**. The current file's mappings are applied to each chosen file (if its sheet names differ, you are asked to match them, with smart suggestions) and each result is saved as a new copy next to that file's original. The other files' own panels are not changed.

### Address syntax

| Input | Meaning |
|---|---|
| `B5` | one cell on the sheet currently shown |
| `Survey!B5` | one cell on the sheet named *Survey* |
| `'My Sheet'!B5:B40` | a range (quotes needed when the name has spaces) |
| `Survey!B5:B10 ; Info!C3` | several ranges, separated by `;` |
| `'[Other.xlsx]Survey'!B5` | a cell in a sheet of another (source) file |

### How sources and targets are paired

| Sources | Targets | Result |
|---|---|---|
| N cells | N cells | paired one-to-one, in order (row by row or column by column) |
| many cells | 1 cell | combined using the chosen mode (first non-empty / join / sum) |
| anything else | | rejected with a message showing the counts |

## When sheet names differ between files

If a template uses sheets that don't exist (by name) in the file you open, the app asks you to match them: for each sheet in the template you choose which sheet of the current file to use, or skip it. A suggestion is pre-selected using, in order: the choice you made last time for that template, similar names, and the sheet's position in the workbook. Your choices are remembered per template, and the template itself keeps its original sheet names, so it still works with the files it was made for. Mappings you skip are marked *Sheet missing* and are not written.

## Where templates are stored

Templates are saved as plain JSON in `~/.excel_report_filler/templates.json`. They contain **cell addresses only, never your data**, so you can copy the file to another computer to share your templates.

## Build a standalone .exe (optional)

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --icon=excel_report_filler.ico --add-data "excel_report_filler.ico;." excel_report_filler.py
```

## Limitations

- The output is written with openpyxl, so content it can't preserve (for example some images, charts or advanced features) may not survive in the saved copy. Always check the result.
- Source values are read as **calculated values**; formulas that were never calculated and saved by Excel will read as empty.
- The sheet preview shows up to 3000 rows and 80 columns per sheet.

## Project structure

```
excel_report_filler.py    # the whole application
excel_report_filler.ico   # app icon
README.md
```

## Author

Developed by **Dev. Marcel Soufi**.