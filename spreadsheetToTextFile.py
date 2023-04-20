#! python3
# spreadsheetToTextFile.py — An exercise in manipulating Excel files.
# For more information, see project_details.txt

import logging
import openpyxl

logging.basicConfig(
    level=logging.DEBUG,
    filename="logging.txt",
    format="%(asctime)s -  %(levelname)s -  %(message)s",
)


def open_workbook():
    workbook = openpyxl.load_workbook("text_to_spread.xlsx")
    sheet = workbook.active
    logging.debug(sheet)


open_workbook()
