from ExcelReader import ExcelReader
from xlrd import open_workbook
from xlutils.copy import copy

import xlrd, xlwt, xlutils, os


class ExcelWriter:

    def __init__(self, target_path, target_name, target_sheet_name='sheet'):
        self.target_path = target_path
        self.target_name = target_name
        self.target_full_path = os.path.join(target_path, target_name)
        self.target_sheet_name = target_sheet_name

    def copy(self, source_path, source_name):
        rb = open_workbook(os.path.join(source_path, source_name), 
                           formatting_info=True, on_demand=True)
        wb = copy(rb)
        wb.save(self.target_full_path)

    def change_cell(self, row, col, value, sheet_num=0):
        rb = open_workbook(os.path.join(self.target_path, self.target_name),
                            formatting_info=True, on_demand=True)
        wb = copy(rb)
        wb.get_sheet(sheet_num).write(row, col, value)
        wb.save(self.target_full_path)

if __name__=="__main__":
    path = "."
    target_name = "测试.xls"
    source_name = "template.xls"

    writer = ExcelWriter(path, target_name)
    writer.copy(path, source_name)

    writer.change_cell(1,1,123)