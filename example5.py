from openpyxl import load_workbook

wb = load_workbook('test2.xlsx')
ws = wb.active

ws.delete_rows(5)
ws.delete_cols(2)

wb.save('test3.xlsx')
