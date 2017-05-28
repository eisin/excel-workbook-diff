import sys
import openpyxl

def excel_workbook_diff(file1, file2, title_row):
    print("file:{} {}".format(file1, file2))
    sheet1 = read_workbook(file1, 0)
    sheet2 = read_workbook(file2, 0)
    
    title_row_cells = []
    for cell in sheet1.iter_cols(min_row=title_row, max_row=title_row):
        title = cell[0].value
        if title is None:
            title = ""
        title.replace("\n", "")
        title_row_cells.append(title)
    sheet2_title_row_cells = []
    for cell in sheet2.iter_cols(min_row=title_row, max_row=title_row):
        title = cell[0].value
        if title is None:
            title = ""
        title.replace("\n", "")
        sheet2_title_row_cells.append(title)
    if title_row_cells != sheet2_title_row_cells:
        raise Exception("title fields not matched. {} vs {}", title_row_cells, sheet2_title_row_cells)
    print (",".join(title_row_cells))

def read_workbook(filename, sheet_index):
    workbook = openpyxl.load_workbook(filename, data_only=True)
    if str(sheet_index).isdigit():
        worksheet = workbook.worksheets[sheet_index]
    else:
        worksheet = workbook.get_sheet_by_name(sheet_index)
    return worksheet

if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise Exception("Invalid arguments")
    file1 = sys.argv[1]
    file2 = sys.argv[2]
    excel_workbook_diff(file1, file2, 1)
