# XlFit

A command-line tool that auto-fits columns and rows in Excel files via COM automation on Windows, with optional integration into the Windows Explorer right-click context menu for GUI use.

---

## Overview

`XlFit` opens a given `.xls` or `.xlsx` file through Excel COM automation, applies auto-fit to every column and row on every sheet, and saves the file in place, all without ever showing the Excel window. Feedback is always given through native Windows dialogs, whether the tool is run from the command line or through the Windows Explorer right-click context menu.

It is particularly useful as a quick finishing step after generating spreadsheets programmatically (e.g. via `openpyxl`, `pandas`, or exported reports), which often leave column widths at their default, unreadable size.

---

## Features

- Auto-fits column width and row height on every sheet of the workbook
- Runs Excel invisibly in the background, no Excel window is shown during processing
- Detects if the target file is already open in another program and warns the user instead of failing silently
- Feedback is always given through native Windows message boxes
- Can be integrated into the Windows Explorer right-click context menu for `.xls`/`.xlsx` files, via the included `.reg` files

---

## Requirements

- Windows with Microsoft Excel installed (uses Excel COM automation)
- Python 3.10+
- Uses standard library modules (`argparse`, `ctypes`, `os`, `sys`) plus the following third-party package: `pywin32`

---

## Installation

```bash
git clone https://github.com/pyquerci/XlFit.git
cd XlFit
```

No installation required. Run directly with Python.

### Windows

Two pre-compiled Windows executables are included in the repository, built with PyInstaller 6.19.0 from the same `XlFit.py` script but with different build options, depending on how you intend to use them:

| Executable | Intended use | Build command |
|---|---|---|
| `XlFit.exe` | Command line: `-h`, `-a`, and `-f` output is shown in the console | `pyinstaller --onefile --manifest="XlFit.manifest" XlFit.py` |
| `XlFitGUI.exe` | Windows Explorer context menu / "Open with": no console window ever shown | `pyinstaller --onefile --noconsole --icon="XlFit.ico" --manifest="XlFit.manifest" XlFit.py` |

Both are built from the exact same code; the only difference is the presence of `--noconsole` and `--icon`. `--manifest` embeds `XlFit.manifest` in both, which declares per-monitor DPI awareness so dialogs render sharply on high-DPI displays.

No Python installation is needed, just download and run the executable you need. For convenience, you can add the folder containing them to your system `PATH`; for example, I keep mine in `C:\Tools\XlFit`.

**Why two executables?** `--noconsole` is required for a clean, flash-free experience when launching from the context menu or "Open with", but it also means `-h/--help` and `-a/--about` produce no visible output, since there is no console attached to print to. That's expected: `XlFitGUI.exe` isn't designed to be used via CLI. Keeping `XlFit.exe` console-enabled means the full CLI, including `-h` and `-a`, always works as expected from a terminal. `XlFit.exe` is meant to be used via CLI only.

#### About the icon

The icon referenced when building `XlFitGUI.exe` (`XlFit.ico`) is included in this repository under the terms described in [Icon Credits](#icon-credits). If you'd rather use your own icon, replace `XlFit.ico` before building `XlFitGUI.exe`.

---

## Usage

```
XlFit.py [-h]
         [-a]
         [-f FILE]
```

### Arguments

| Argument | Description |
|---|---|
| `-h, --help` | Show the help message and exit. |
| `-a, --about` | Show author, version, project URL and license information. |
| `-f, --file FILE` | Excel file to auto-fit. |

### Examples

```bash
# Auto-fit a single Excel file
XlFit.exe -f report.xlsx

# Show author and version information
XlFit.exe -a
```

If the file is already open in another program, a warning dialog is shown and the file is left untouched. On success, a confirmation dialog is displayed; on failure, an error dialog reports what went wrong.

---

## How It Works

1. `XlFit` checks whether the target file is currently locked by another process (e.g. already open in Excel). If so, a warning dialog is shown and execution stops.
2. Excel is launched in the background via COM automation (`Excel.Application`), with its window hidden and alerts disabled.
3. The workbook is opened and for every sheet the used range is calculated, column widths are set to a wide baseline, then auto-fit is applied to both columns and rows.
4. The workbook is saved and closed.
5. A native Windows dialog reports success or, if an error occurred, what went wrong.

---

## Windows Explorer Context Menu Integration

Two registry files are included for this purpose:

- `AddContextMenu.reg` — adds an **"Excel Auto Fit"** entry to the right-click context menu of `.xls`/`.xlsx` files, and registers `XlFitGUI.exe` in the "Open with" list for both extensions
- `RemoveContextMenu.reg` — removes both integrations

### Installing the context menu entry

1. Place `XlFitGUI.exe` in the path referenced by `AddContextMenu.reg` (`C:\Tools\XlFit\XlFitGUI.exe` by default), or edit the `.reg` file to match your own installation path.
2. Double-click `AddContextMenu.reg` and confirm the prompt to merge it into the registry.

Once installed, right-clicking any `.xls`/`.xlsx` file in Explorer will show an **"Excel Auto Fit"** option, and `XlFitGUI.exe` will also appear as an option in the "Open with" dialog.

### Removing the context menu entry

Double-click `RemoveContextMenu.reg` and confirm the prompt. This deletes the registry keys created by `AddContextMenu.reg`, removing both the context menu entry and the "Open with" registration.

---

### Example

The GIF below shows both ways to use `XlFit` from Windows Explorer in Windows 11:

![Excel Auto Fit demo](XlFitGUI.gif)

---

## Known Issues

- Although Excel is launched with `Visible = False`, the Excel window may occasionally flash briefly on screen during processing. This is caused by Excel's own internal behavior during COM automation and is not something `XlFit` can fully control; it happens unpredictably and does not affect the outcome of the operation.
- `XlFitGUI.exe` is built with `--noconsole`, so `-h/--help` and `-a/--about` produce no visible output, even when run from a terminal, since the executable has no console attached to print to. Use `XlFit.exe` instead if you need the full CLI, including help and about.

---

## Icon Credits

XlFit.ico icon used in this project was downloaded from `Icon-Icons.com` under the CC BY 4.0 license.

- Author: `Those Icons`.
- Source: https://icon-icons.com/authors/639-those-icons

---

## License

This project is licensed under the **GNU General Public License v2.0 (GPLv2)**. You are free to use, modify, and distribute this software under the terms of that license. See the [LICENSE](LICENSE) file for the full license text.

---

## Donations

If you value the work and want to help support its development, feel free to make a donation. Your support will be greatly appreciated:

- PayPal: https://paypal.me/pyquerci
- Buy Me a Coffee: https://buymeacoffee.com/pyquerci
