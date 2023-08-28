#!/usr/bin/env python3
# spreadsheetToTextFile.py — An exercise in manipulating Excel files.
# For more information, see README.md

import logging
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.DEBUG,
    filename="logging.txt",
    format="%(asctime)s -  %(levelname)s -  %(message)s",
)
logging.disable(logging.CRITICAL)  # Note out to enable logging.


def main() -> None:
    """Runs functions in sequence on user-designated file."""
    file_name = input("Please enter file name here: ")
    file_contents = record_contents(file_name)
    write_contents(file_contents)


def record_contents(file: str) -> list[list[str]]:
    """Open Excel file in current directory and record contents of
    rows per column of sheet into a list of lists, then return list."""
    sheet = load_workbook(f"{file}.xlsx").active
    return [
        [
            sheet[f"{get_column_letter(column + 1)}{row}"].value
            for row in range(1, sheet.max_row + 1)
            if sheet[f"{get_column_letter(column + 1)}{row}"].value is not None
        ]
        for column in range(sheet.max_column)
    ]


def write_contents(contents: str) -> None:
    """Write contents of each line of secondary list to file associated with primary list value."""
    for i, _ in enumerate(contents):
        with open(f"text_of_column{(i + 1):03}.txt", "w", encoding="utf-8") as txt:
            for j, _ in enumerate(contents[i]):
                txt.write(str(contents[i][j]))


if __name__ == "__main__":
    main()
