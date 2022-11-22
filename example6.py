from openpyxl import load_workbook

wb = load_workbook('test3.xlsx')
ws = wb.active

ws.move_range("C1:D7", rows=2, cols=5)

wb.save('test4.xlsx')
