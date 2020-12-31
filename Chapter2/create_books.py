from openpyxl import Workbook

count = input('作成するブック数 : ')
for i in range(int(count)):
    # 変数wbにブックを代入(※この段階ではファイルとしては保存されていない)
    wb = Workbook()
    """ 宣言したブックのオブジェクトに対して、.active属性を指定すると、
    ブック作成時、デフォルトで１つ作成されるシート(=アクティブなシート)の名前を
    取得出来る(デフォルトは「Sheet」)
     """
    ws = wb.active
    """ デフォルトのシート名を格納したWorkSheetのオブジェクトに対して、
    title属性でシート名を指定することになり、WorkSheetのシート名に「概要」を代入
     """
    ws.title = '概要'
    # WorkBookオブジェクトに対してsaveメソッドを実行することで、ブックが保存される
    # 引数には、保存時に付けたいファイル名を記述
    # 「f''」→「フォーマット済み文字列リテラル」・・・ある文字列内に変数の値を含める際に使用
    # 今回の場合は、saveメソッドの引数にしている、ファイル名指定時に使用
    # ループ変数iは0から始まる為、1を加算。
    wb.save(f'資料_{i + 1}.xlsx')
