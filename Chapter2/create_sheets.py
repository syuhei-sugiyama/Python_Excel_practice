from openpyxl import Workbook

count = input('全シート数 : ')

# Workbookオブジェクト生成
wb = Workbook()
# WorkSheetオブジェクト取得
ws = wb.active
# デフォルトで作成されるシート名を変更
ws.title = '概要_1'

for i in range(2, int(count) + 1):
    # シートを作成する
    # シート名はtitle引数で指定する
    # f-stringsを使用
    wb.create_sheet(title=f'概要_{i}')

wb.save('シート数指定.xlsx')
