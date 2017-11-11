import os.path
from openpyxl import Workbook, load_workbook


def write(fields, ldap_q_results, sn):
    xlsfile = 'firedrill.xlsx'
    if not os.path.exists(xlsfile):
        wb = Workbook(write_only=True)
    else:
        wb = load_workbook(filename = xlsfile)
    try:
        ws = wb[sn]
    except:
        ws = wb.create_sheet(sn)
    
    ws.append(fields)
    for result in ldap_q_results:
        ws.append([result[k] for k in fields])
    wb.save(xlsfile)


if __name__ == '__main__':
    # test code
    fields = ['uid', 'uidNumber']
    results = [{'uid': 'haitchen', 'uidNumber': '61418924'},
               {'uid': 'haizzhang', 'uidNumber': '61418024'}]
    write(fields, results) 
