""" Application to analyze an .xlsx file"""

import os
import sys
import openpyxl
from spreadsheet import Spreadsheet

workbook = Spreadsheet()
for sheet in workbook.sheets:
    for key, value in sheet.items():
        print("{}:\t{}".format(key, value))