from openpyxl import load_workbook

wb = load_workbook('test.xlsx')
ws = wb.active

ws.insert_rows(5)
ws.insert_cols(2)

wb.save('test2.xlsx')
