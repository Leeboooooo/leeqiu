

from openpyxl import Workbook
from openpyxl import load_workbook

# wb = Workbook()

# ws = wb.active
# fileOne = load_workbook('D:\\ghosttusng\\RDEXcel\\1.xlsx')

import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )

def readExcel():
    fileName = r'D:/ghosttusng/RDExcel/1.xlsx'
    print(fileName)
    # inwb = openpyxl.load_workbook(fileName)
    wb = load_workbook(fileName)
    # print(wb.sheetnames)
    ws = wb.active
    for sheet in wb:
        print(sheet.title)
    

    # ws = inwb.get_sheet_by_name(sheetNames[0])
    rows = ws.max_row
    cols = ws.max_column
    for r in range(1, rows):
            for c in range(1, cols):
                print(ws.cell(r, c).value)
            # if r == 10:
            #     break
    return;

readExcel();