import sys
import csv
import RegistrationInfo
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

def read_regs(csvfile='Form_data.csv'):
    users = []
    with open(csvfile, 'rb') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            try:
                int(row[0])
            except:
                print 'skip the fields'
                continue
            if 'employee' in row[1]:
                print 'this is Employee, registered @ %s'%(row[6])
                user = RegistrationInfo.RegistrationInfo('nsb', row[6], uidnum=row[3], wechat_nickname=row[8].decode('gbk').encode('utf-8'))
            else:
                print 'this is visitor, registered @ %s'%(row[6])
                user = RegistrationInfo.RegistrationInfo('ext', row[6], company=row[2], wechat_nickname=row[8].decode('gbk').encode('utf-8'))
            users.append(user)
    return users
            #for cell in row:
            #    print cell.decode('gbk')

__all__ = [read_regs]

if __name__ == '__main__':
    #print(read_employees())
    allusers = read_regs()
    for user in allusers:
        print user

