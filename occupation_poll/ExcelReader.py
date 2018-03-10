import xlrd, os

class ExcelReader:
    
    def __init__(self, filepath, filename):
        self.rtable = self.__load_sheet(filepath, filename)
        pass

    def __load_sheet(self, filepath, filename):
        fullpath = os.path.join('%s/%s' % (filepath, filename))
        source = xlrd.open_workbook(fullpath)
        rtable = source.sheets()[0]
        return rtable

    def read_cell(self, row, col):
        return self.rtable.cell(row, col).value

    def read_column(self, row_range, col):
        result = []
        for row in row_range:
            result.append(self.read_cell(row, col))
        return result

    def read_two_column_as_pair(self, row_range, col1, col2):
        result = {}
        for row in row_range:
            result[self.read_cell(row, col1)] = self.read_cell(row, col2)
        return result

if __name__ == "__main__":
    filepath = "."
    filename = "2015master.xls"
    er = ExcelReader(filepath, filename)

    classmates_range = range(3, 90)
    num_name = er.read_two_column_as_pair(classmates_range, 1, 2)
    print(num_name)
    print(len(num_name))
