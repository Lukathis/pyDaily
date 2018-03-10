from ExcelReader import ExcelReader
from ExcelWriter import ExcelWriter

if __name__ == "__main__":
    filepath = "."
    filename = "2015master.xls"
    er = ExcelReader(filepath, filename)

    classmates_range = range(3, 91)
    # classmates_range = [90]
    num_name = er.read_two_column_as_pair(classmates_range, 1, 2)
    
    source_path = "."
    source_name = "template.xls"
    target_path = "./blank"
    

    for num in num_name:
        name = num_name[num]
        target_name = "{}.xls".format(num, name)
        writer = ExcelWriter(target_path, target_name)
        writer.copy(source_path, source_name)
        writer.change_cell(2, 1, num)
        writer.change_cell(2, 2, name)



