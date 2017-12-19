#-*- encoding=utf-8 -*-

import sys
import csv
import re
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

def in_valid_area(loc):
    pass

def read_regs(csvfile='Form_data.csv'):
    nicknames = []
    users = []
    with open(csvfile, 'rb') as csvfile:
        reader = csv.reader(csvfile)
        for row in reversed(list(reader)):
            try:
                int(row[0])
            except:
                print 'skip the fields'
                continue
            nickname = row[9].decode('gbk').encode('utf-8')
            location = row[4].decode('gbk').encode('utf-8')
            if nickname in nicknames:
                continue
            nicknames.append(nickname)
            if 'employee' in row[1]:
                print 'this is Employee, registered @ %s'%(row[7])
                user = RegistrationInfo.RegistrationInfo('nsb', row[7], uidnum=row[3], 
                    wechat_nickname=nickname,
                    location=location)
            else:
                print 'this is visitor, registered @ %s'%(row[7])
                user = RegistrationInfo.RegistrationInfo('ext', row[7], company=row[2],
                    wechat_nickname=nickname,
                    location=location)
            if user.is_valid_location():
                users.append(user)
    users.reverse()
    return users
            #for cell in row:
            #    print cell.decode('gbk')

__all__ = [read_regs]

if __name__ == '__main__':
    #print(read_employees())
    allusers = read_regs()
    for user in allusers:
        print user

