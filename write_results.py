import os.path
import xlsxwriter
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
    if sn == 'management tree sheet':
        tm_wb = xlsxwriter.Workbook('fd_mngm_tr.xlsx')
        ws = tm_wb.add_worksheet('management tree')
        row = 0
        for col, f in enumerate(fields):
            ws.write(row, col, f)
        row += 1
        topmgr = ''
        topsub = 0
        mgr2 = ''
        sub2 = 0
        mgr3 = ''
        sub3 = 0
        mgr4 = ''
        sub4 = 0
        pretop = ''
        pre2 = ''
        pre3 = None
        pre4 = None
        levels = 0
        bold = tm_wb.add_format({'bold': 1})
        for d in ldap_q_results:
            topmgr = d['manager-1']
            mgr2 = d['manager-2']
            mgr3 = d['manager-3']
            mgr4 = d['manager-4']
            if pretop == '':
                pretop = topmgr
            if pre2 == '':
                pre2 = mgr2
            if pre3 is None:
                pre3 = mgr3
            if pre4 is None:
                pre4 = mgr4
            if pre4 != mgr4:
                if levels > 4 or (levels == 4 and sub4 == 1):
                    ws.write('D%d'%(row+1,), pre4, bold)
                    ws.write('E%d'%(row+1,), 'Expected No.: ', bold)
                    ws.write('F%d'%(row+1,), 'Actual No.: ', bold)
                    ws.set_row(row, None, None, {'level': 4, 'hidden': True, 'collapsed': True})
                    row += 1
                sub4 = 0
            else:
                sub4 += 1
            if pre3 != mgr3:
                if levels > 3 or (levels == 3 and sub3 == 1):
                    ws.write('C%d'%(row+1,), pre3, bold)
                    ws.write('D%d'%(row+1,), 'Expected No.: ', bold)
                    ws.write('E%d'%(row+1,), 'Actual No.: ', bold)
                    ws.set_row(row, None, None, {'level': 3, 'hidden': True, 'collapsed': True})
                    row += 1
                sub3 = 0 # clear sub counter
            else:
                sub3 += 1
            if pre2 != mgr2:
                ws.write('B%d'%(row+1,), pre2, bold)
                ws.write('C%d'%(row+1,), 'Expected No.: ', bold)
                ws.write('D%d'%(row+1,), 'Actual No.: ', bold)
                ws.set_row(row, None, None, {'level': 2, 'hidden': True, 'collapsed': True})
                row += 1
            else:
                sub2 += 1
            if pretop != topmgr:
                ws.write('A%d'%(row+1,), pretop, bold)
                ws.write('B%d'%(row+1,), 'Expected No.: ', bold)
                ws.write('C%d'%(row+1,), 'Actual No.: ', bold)
                ws.set_row(row, None, None, {'level': 1, 'hidden': True, 'collapsed': True})
                row += 1
            else:
                topsub += 1
            for col, f in enumerate(fields):
                ws.write(row, col, d[f])
                ws.set_row(row, None, None, {'level': 5, 'hidden': True})
            pretop = topmgr
            pre2 = mgr2
            pre3 = mgr3
            if sub3 == 0:
                sub3 += 1
            levels = len(d['managers'])
            pre4 = mgr4
            if sub4 == 0:
                sub4 += 1
            row += 1
        row += 1
        ws.set_row(row, None, None, {'collapsed': False})
        tm_wb.close()
        
    wb.save(xlsfile)

def group(xlsfile, sheet, row, level):
    wb = xlsxwriter.Workbook(xlsfile)
    ws = wb.load_worksheet(sheet)


if __name__ == '__main__':
    # test code
    fields = ['uid', 'uidNumber']
    results = [{'uid': 'haitchen', 'uidNumber': '61418924'},
               {'uid': 'haizzhang', 'uidNumber': '61418024'}]
    write(fields, results) 
