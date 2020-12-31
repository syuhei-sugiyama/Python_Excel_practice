from openpyxl import Workbook

# Excelファイルを新規作成
wb = Workbook()
ws = wb.active

for row_count in range(1, 5):
    print(row_count)
    cell_no = f'A{row_count}'
    ws[cell_no] = 'Hello'

wb.save('test.xlsx')
