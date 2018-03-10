import xlrd, xlwt, os

class AssembleClient:

    def __init__(self, source_path, source_row, column_range, 
                 target_file_name="result", target_sheet_name="sheet"):
        SUFFIX = ".xls"
        self.source_path = source_path
        self.source_row = source_row
        self.column_range = column_range
        self.target_file_name = target_file_name + SUFFIX
        self.target_sheet_name = target_sheet_name

    def assemble(self):
        target =  xlwt.Workbook()
        wtable = target.add_sheet(self.target_sheet_name, cell_overwrite_ok=True)
        
        index = 0
        for filename in self.get_file_list(self.source_path):
            rtable = self.read_content(self.source_path, filename)
            self.write_content(wtable, index, rtable)
            index += 1
            
        target.save(self.target_file_name)
        print("Completed!")

    def get_file_list(self, path):
        childs = os.listdir(path)

        result = []
        for a in childs:
            if a.endswith(".xlsx") or a.endswith(".xls"):
                result.append(a)
        return result

    def read_content(self, filepath, filename):
        fullpath = os.path.join('%s/%s' % (filepath, filename))
        source = xlrd.open_workbook(fullpath)
        rtable = source.sheets()[0]
        return rtable
        
    def write_content(self, wtable, target_row, rtable):

        if isinstance(self.column_range, int):
            column_range = list(range(0, self.column_range))
        elif isinstance(self.column_range, list):
            column_range = self.column_range

        target_index = 0
        for i in column_range:
            wtable.write(target_row, target_index, rtable.cell(self.source_row, i).value)
            target_index += 1
        return wtable

if __name__ == "__main__":

    SOURCE_PATH = './blank'
    SOURCE_ROW = 2
    COLUMN_RANGE = 5

    ac = AssembleClient(SOURCE_PATH, SOURCE_ROW, COLUMN_RANGE, "result1", "sheet")
    ac.assemble()

    


