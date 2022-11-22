# https://openpyxl.readthedocs.io/en/stable/
from openpyxl import Workbook


wb = Workbook()
ws = wb.active
ws.title = 'Data'
ws.append(['№', 'My', 'first', 'project', '!'])
for i in range(15):
    ws.append([i + 1, 'My', 'first', 'project', '!'])

wb.save('test.xlsx')
