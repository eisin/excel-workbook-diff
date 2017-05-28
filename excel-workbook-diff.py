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
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            pass
        elif tag == "insert":
            for i in range(j1, j2):
                print("insert #{}:{}".format(i, table2[i]))
        elif tag == "delete":
            for i in range(i1, i2):
                print("delete #{}:{}".format(i, table1[i]))
        elif tag == "replace":
            insert_entries = list(table2[j1:j2])
            for i in range(i1, i2):
                while True:
                    replace_one_line_score = 0
                    if len(insert_entries) >= 1:
                        replace_one_line_score = count_exact_entries_in_tuple(table1[i], insert_entries[0])
                    insert_one_line_score = 0
                    if len(insert_entries) >= 2:
                        insert_one_line_score = count_exact_entries_in_tuple(table1[i], insert_entries[1])
                    delete_one_line_score = 0
                    if i < i2 and len(insert_entries) >= 1:
                        delete_one_line_score = count_exact_entries_in_tuple(table1[i + 1], insert_entries[0])
                    print("SCORE replace({}) insert({}) delete({})".format(replace_one_line_score, insert_one_line_score, delete_one_line_score))
                    if max(replace_one_line_score, insert_one_line_score, delete_one_line_score) == replace_one_line_score:
                        print("change(before) #{}:{}".format(i, table1[i]))
                        print("change(after) #{}:{}".format(i, insert_entries.pop(0)))
                        break
                    elif max(replace_one_line_score, insert_one_line_score, delete_one_line_score) == insert_one_line_score:
                        print("insert #{}:{}".format(i, insert_entries.pop(0)))
                    else:
                        print("delete #{}:{}".format(i, table1[i]))
                        break
            
        
def count_exact_entries_in_tuple(tuple1, tuple2):
    count = 0
    for i in range(min(len(tuple1), len(tuple2))):
        if tuple1[i] == tuple2[i]:
            count += 1
    return count

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
