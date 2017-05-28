import sys
import openpyxl
from difflib import SequenceMatcher

def diff_excel_workbook(file1, file2, title_row):
    print("file:{} {}".format(file1, file2))
    sheet1 = read_workbook(file1, 0)
    sheet2 = read_workbook(file2, 0)
    
    sheet1_header_titles = []
    sheet2_header_titles = []
    for cells in sheet1.iter_cols(min_row=title_row, max_row=title_row):
        sheet1_header_titles.append(cell_to_text_oneline(cells[0]))
    for cells in sheet2.iter_cols(min_row=title_row, max_row=title_row):
        sheet2_header_titles.append(cell_to_text_oneline(cells[0]))
    if sheet1_header_titles != sheet2_header_titles:
        raise Exception("title fields not matched. {} vs {}", sheet1_header_titles, sheet2_header_titles)

    table1 = read_sheet_table(sheet1)
    table2 = read_sheet_table(sheet2)
    diff_result = diff_two_tables(table1, table2)

    for opcode, field1, field2 in diff_result:
        if opcode == "insert":
            print("INSERT:" + ",".join(field1))
        elif opcode == "delete":
            print("DELETE:" + ",".join(field1))
        elif opcode == "replace":
            print("DELREP:" + ",".join(field1))
            print("INSREP:" + ",".join(field2))
        else:
            raise Exception("Could not recognize diff opcode:{}".format(opcode))

def diff_two_tables(table1, table2):
    matcher = SequenceMatcher(None, table1, table2)
    result = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            pass
        elif tag == "insert":
            for i in range(j1, j2):
                result.append(("insert", table2[i],None,))
        elif tag == "delete":
            for i in range(i1, i2):
                result.append(("delete", table1[i],None,))
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
                        #print("change(before) #{}:{}".format(i, table1[i]))
                        #print("change(after) #{}:{}".format(i, insert_entries.pop(0)))
                        result.append(("replace", table1[i], insert_entries.pop(0)))
                        break
                    elif max(replace_one_line_score, insert_one_line_score, delete_one_line_score) == insert_one_line_score:
                        result.append(("insert", table2[i],None,))
                    else:
                        result.append(("delete", table1[i],None,))
                        break
    return result
        
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
    diff_excel_workbook(file1, file2, 1)
