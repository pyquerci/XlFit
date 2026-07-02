# XlFit

A command-line tool that auto-fits columns and rows in Excel files via COM automation on Windows, with optional integration into the Windows Explorer right-click context menu for GUI use.

---

## Overview

`XlFit` opens a given `.xls`or `.xlsx` file through Excel COM automation, applies auto-fit to every column and row on every sheet, and saves the file in place, all without ever showing the Excel window. Feedback is always given through native Windows dialogs, whether the tool is run from the command line or through the Windows Explorer right-click context menu.

It is particularly useful as a quick finishing step after generating spreadsheets programmatically (e.g. via `openpyxl`, `pandas`, or exported reports), which often leave column widths at their default, unreadable size.

---

## Features

- Auto-fits column width and row height on every sheet of the workbook
- Runs Excel invisibly in the background, no Excel window is shown during processing
- Detects if the target file is already open in another program and warns the user instead of failing silently
- Feedback is always given through native Windows message boxes
- Can be integrated into the Windows Explorer right-click context menu for `.xlsx` files, via the included `.reg` files

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

A pre-compiled Windows executable is included in the repository, built with PyInstaller 6.19.0 using the command:

```bash
pyinstaller --onefile --noconsole --icon="XlFit.ico" --manifest="XlFit.manifest" XlFit.py
```

- `--noconsole` suppresses the console window, since all feedback is given through native message boxes.
- `--icon` sets the custom icon shown by the executable and, where configured, by the context menu entry.
- `--manifest` embeds `XlFit.manifest`, which declares per-monitor DPI awareness so dialogs render sharply on high-DPI displays.

No Python installation is needed, just download and run `XlFit.exe`. For convenience, you can add it to a folder in your system `PATH` to invoke it from any directory; for example, I keep mine in `C:\Tools\XlFit`.

**Note:** `--noconsole` is recommended if you only intend to use `XlFit.exe` through the `-f/--file` argument, either from the command line or from the Windows Explorer context menu, since feedback is always given via native dialogs regardless of the console. However, this means `-h/--help` and `-a/--about` will produce no visible output when run from the compiled `.exe`, since there is no console attached to print to (see [Known Issues](#known-issues)). If you want a fully functional CLI, including `-h` and `-a`, rebuild without `--noconsole`:

```bash
pyinstaller --onefile --icon="XlFit.ico" --manifest="XlFit.manifest" XlFit.py
```

#### About the icon

The icon referenced during the build (`XlFit.ico`) is **not included** in this repository, since it was originally extracted from Microsoft Excel's own resources — which are copyrighted and cannot be redistributed.

To use your own icon:

- Extract an icon from `excel.exe` yourself using a free tool such as [Resource Hacker](http://www.angusj.com/resourcehacker/) or [IcoFX](https://icofx.ro/), or
- Use any `.ico` file of your choice (e.g. from a free icon set like [Icons8](https://icons8.com) or [Flaticon](https://www.flaticon.com))

Then reference your chosen `.ico` file both in the `pyinstaller --icon` argument and, if you set up the context menu integration, in `AddContextMenu.reg`.

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
XlFit.py -f report.xlsx

# Show author and version information
XlFit.py -a
```

If the file is already open in another program, a warning dialog is shown and the file is left untouched. On success, a confirmation dialog is displayed; on failure, an error dialog reports what went wrong.

---

## How It Works

1. `XlFit` checks whether the target file is currently locked by another process (e.g. already open in Excel). If so, a warning dialog is shown and execution stops.
2. Excel is launched in the background via COM automation (`Excel.Application`), with its window hidden and alerts disabled.
3. The workbook is opened, and for every sheet: the used range is calculated, column widths are set to a wide baseline, then auto-fit is applied to both columns and rows.
4. The workbook is saved and closed.
5. A native Windows dialog reports success or, if an error occurred, what went wrong.

---

## Windows Explorer Context Menu Integration

`XlFit` accepts a single file via `-f/--file`, which makes it a natural fit for the Windows Explorer right-click context menu: right-clicking an `.xlsx` file passes its path straight to `XlFit.exe -f "%1"`.

Two registry files are included for this purpose:

- `AddContextMenu.reg` — adds an **"Excel Auto Fit"** entry to the right-click context menu of `.xlsx` files
- `RemoveContextMenu.reg` — removes it

### Installing the context menu entry

1. Place `XlFit.exe` in the path referenced by `AddContextMenu.reg` (`C:\Tools\XlFit\XlFit.exe` by default), or edit the `.reg` file to match your own installation path and icon.
2. Double-click `AddContextMenu.reg` and confirm the prompt to merge it into the registry.

Once installed, right-clicking any `.xlsx` file in Explorer will show an **"Excel Auto Fit"** option that runs `XlFit.exe -f` directly on that file.

### Removing the context menu entry

Double-click `RemoveContextMenu.reg` and confirm the prompt. This deletes the registry key created by `AddContextMenu.reg`, removing the **"Excel Auto Fit"** entry from the context menu.

---

## Icon Credits

XlFit.ico icon used in this project was downloaded from Icon-Icons.com under the CC BY 4.0 license.

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
