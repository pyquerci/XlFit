# Author: Andrea Querci
# version: 1.0.0
# project: https://github.com/pyquerci/XlFit
# license: GPLv2

import argparse
import ctypes
import os
import sys
from win32com.client import Dispatch


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def is_file_locked(file_path: str) -> bool:
    try:
        with open(file_path, "r+b"):
            return False
    except (IOError, OSError):
        return True


def excel_auto_fit(excel, file_path: str) -> None:
    wb = excel.Workbooks.Open(file_path)
    try:
        for ws in wb.Sheets:
            ws.Activate()
            used = ws.UsedRange
            if used is None:
                continue
            ncols = used.Column + used.Columns.Count - 1
            last_col = colnum_string(ncols)
            ws.Range(f"A:{last_col}").ColumnWidth = 100
            ws.Columns.AutoFit()
            ws.Rows.AutoFit()
        excel.Worksheets(1).Activate()
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        usage=("%(prog)s [-h] "
            "[-a] "
            "[-f FILE]"
        ),
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=(
            "description:\n"
            "  auto-fit columns and rows in an Excel file via COM automation\n"
            "  XlFit can also be integrated in the Windows Explorer right-click context menu\n"
            "  see the project documentation for more info\n\n"
        ),
    )

    parser.add_argument(
        "-a",
        "--about",
        action="store_true",
        help="show author, version, project URL and license information",
    )

    parser.add_argument(
        "-f",
        "--file",
        metavar="FILE",
        help="excel file to auto-fit",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_arguments()

    if args.about:
        print(
            "author: Andrea Querci\n"
            "version: 1.0\n"
            "project: https://github.com/pyquerci/XlFit\n"
            "license: GPLv2"
        )
        return

    if not args.file:
        print("error: no file specified, use -f/--file")
        sys.exit(1)

    file = os.path.abspath(args.file)
    name = os.path.basename(file)

    if is_file_locked(file):
        ctypes.windll.user32.MessageBoxW(
            0,
            f"Impossibile procedere:\n\n• {name} è attualmente aperto in un altro programma.\n\nChiudi il file e riprova.",
            "Excel Auto Fit",
            0x30,  # MB_ICONWARNING
        )
        sys.exit(1)

    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        excel_auto_fit(excel, file)
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f"Error for file '{name}':\n\n{e}", "Excel Auto Fit", 0x30)
        sys.exit(1)
    finally:
        excel.Quit()

    ctypes.windll.user32.MessageBoxW(0, f"'{name}' fixed successfully", "Excel Auto Fit", 0x40)


if __name__ == "__main__":
    main()
