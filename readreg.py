#-*- encoding=utf-8 -*-

import sys
import csv
import re
import chardet
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
    unknown_encoding = 0
    with open(csvfile, 'rb') as csvfile:
        reader = csv.reader(csvfile)
        for row in reversed(list(reader)):
            try:
                int(row[0])
            except:
                print 'skip the fields'
                continue
            try:
                nickname = row[6].decode('gbk').encode('utf-8').strip()
            except:
                try:
                    nickname = row[6].decode('unicode_escape').encode('utf-8').strip()
                except:
                    nickname = 'unknown encoding'
                    unknown_encoding += 1
                    #print nickname
                    print chardet.detect(nickname)
                    #return
            location = row[4].decode('gbk').encode('utf-8')
            if nickname in nicknames:
                continue
            nicknames.append(nickname)
            if 'employee' in row[1]:
                #print 'this is Employee, registered @ %s'%(row[5])
                user = RegistrationInfo.RegistrationInfo('nsb', row[5], uidnum=row[3].encode('utf-8'), 
                    wechat_nickname=nickname,
                    location=location)
            else:
                #print 'this is visitor, registered @ %s'%(row[5])
                user = RegistrationInfo.RegistrationInfo('ext', row[5], company=row[2].decode('gbk').encode('utf-8').strip(),
                    wechat_nickname=nickname,
                    location=location)
            #if user.is_valid_location() and user.is_valid_time():
            if not user.is_valid_location():
                print '%s has wrong location, %s'%(user.nickname, user.location)
            if not user.is_valid_time():
                print '%s has wrong regtime, %s'%(user.nickname, user.regtime)
            users.append(user)
    print 'unknown_encoding: %d'%(unknown_encoding)
    users.reverse()
    return users
            #for cell in row:
            #    print cell.decode('gbk')

__all__ = [read_regs]

if __name__ == '__main__':
    #print(read_employees())
    allusers = read_regs()
    num_of_regs = 0
    for user in allusers:
        if user.is_valid_time and user.is_valid_location:
            num_of_regs += 1
    print num_of_regs


