import sys
import openpyxl
from difflib import SequenceMatcher

def excel_workbook_diff(file1, file2, title_row):
    print("file:{} {}".format(file1, file2))
    sheet1 = read_workbook(file1, 0)
    sheet2 = read_workbook(file2, 0)
    
    title_row_cells = []
    sheet2_title_row_cells = []
    for cells in sheet1.iter_cols(min_row=title_row, max_row=title_row):
        title_row_cells.append(cell_to_text_oneline(cells[0]))
    for cells in sheet2.iter_cols(min_row=title_row, max_row=title_row):
        sheet2_title_row_cells.append(cell_to_text_oneline(cells[0]))
    if title_row_cells != sheet2_title_row_cells:
        raise Exception("title fields not matched. {} vs {}", title_row_cells, sheet2_title_row_cells)
    print (",".join(title_row_cells))

    table1 = read_sheet_table(sheet1)
    table2 = read_sheet_table(sheet2)

    matcher = SequenceMatcher(None, table1, table2)
    #print(matcher)
    #print(list(matcher.get_opcodes()))
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        print(tag, i1, i2, j1, j2)
        if tag == "equal":
            pass
        elif tag == "insert":
            for i in range(j1, j2):
                print("insert #{}:{}".format(i, table2[i]))
        elif tag == "delete":
            for i in range(i1, i2):
                print("delete #{}:{}".format(i, table1[i]))
        elif tag == "replace":
            #for i in range(i1, i2):
            #    print("delete #{}:{}".format(i, table1[i]))
            #for i in range(j1, j2):
            #    print("insert #{}:{}".format(i, table2[i]))
            replace_list = []
            for i in range(i1, i2):
                replace_list.append(table1[i] + ("00-delete",))
            for j in range(j1, j2):
                replace_list.append(table2[j] + ("11-append",))
            replace_list = sorted(replace_list)
            for entry in replace_list:
                action = entry[-1]
                line = entry[0:-1]
                print("{} :{}".format(action, line))
        

def read_workbook(filename, sheet_index):
    workbook = openpyxl.load_workbook(filename, data_only=True)
    if str(sheet_index).isdigit():
        worksheet = workbook.worksheets[sheet_index]
    else:
        worksheet = workbook.get_sheet_by_name(sheet_index)
    return worksheet

def read_sheet_table(sheet, start_row=2):
    row_index = 0
    table = []
    for row in sheet.rows:
        row_index += 1
        if row_index <= start_row:
            continue
        line = []
        for cell in row:
            text = cell.value
            if text is None:
                text = ""
            line.append(str(text))
        table.append(tuple(line))
    return table

def cell_to_text_oneline(cell):
    text = cell.value
    if text is None:
        text = ""
    text.replace("\n", "")
    return text

if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise Exception("Invalid arguments")
    file1 = sys.argv[1]
    file2 = sys.argv[2]
    excel_workbook_diff(file1, file2, 1)
