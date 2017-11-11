import sys
from openpyxl import load_workbook

def read_employees(xlsx_file='tthc930.xlsx'):
    nsn_ids = []
    wb = load_workbook(filename=xlsx_file, read_only=True)
    ws = wb['Sheet1']
    num_of_employees = 0
    for row in ws.rows:
        is_head = False 
        for i, cell in enumerate(row):
            if 'Company' in cell.value:
                is_head = True
                break
            if i == 2:
                nsn_ids.append(cell.value)
        if is_head: continue
        num_of_employees = num_of_employees + 1

    print(num_of_employees)
    return nsn_ids


if __name__ == '__main__':
    print(read_employees())

