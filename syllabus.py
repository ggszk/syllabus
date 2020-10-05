#
# シラバスのExcelファイルの変換
# 

import openpyxl

# ExcelからMapにあるセルを読んで，セルのリストを返す
def read_cells(in_sheet, in_map, offset):
    in_cells = []
    i = 0
    for pt in in_map :
        in_cells.append(in_sheet.cell(row=pt[0]+offset[0], column=pt[1]+offset[1]))
        i = i + 1
    return in_cells

# セルのリストをMapに従ってExcelに書き込む
def write_cells(out_sheet, in_cells, out_map, offset):
    i = 0
    for cell in in_cells :
        out_sheet.cell(row=out_map[i][0]+offset[0], column=out_map[i][1]+offset[1], value=cell.value)
        i = i + 1
    return

# 入力ファイルを出力ファイルにマップの指定に従って書き込む
def syllabus(in_filename, out_filename, in_map_list, out_map_list):
    # ワークブック
    in_wb = openpyxl.load_workbook(in_filename)
    # 出力ファイルは新規作成（テンプレートをベースに）
    out_wb = openpyxl.load_workbook('template.xlsx')

    # 出力テンプレートシート
    temp_sheet = out_wb['template']

    # セルの読み書き
    for in_sheet in in_wb :
        # シート名が科目名
        kamoku = in_sheet.title
        # テンプレートシートをコピーしてリネーム
        out_sheet = out_wb.copy_worksheet(temp_sheet)
        out_sheet.title = kamoku

        # mapは講義回数に応じて選択する
        map_index = 0
        in_map = in_map_list[map_index]
        out_map = out_map_list[map_index]

        # オフセットは，「授業科目」が左上(1,1)の位置からいくつずれているか（ずれなしは(0,0)）
        # セルを読む：入力
        in_cells = read_cells(in_sheet, in_map, (1,0))

        # セルに代入
        # 新
        write_cells(out_sheet, in_cells, out_map, (2,1))
        # 旧
        write_cells(out_sheet, in_cells, out_map, (2,8))

    # ファイル保存（ここでファイルは新規作成される）
    out_wb.save(out_filename)

# 入出力ファイル
#in_filename = "./temp/シラバスまとめ　情報（教育課程等の概要の順番） 59.xlsx"
#in_filename2 = "./temp/シラバスまとめ　事業創造（教育課程等の概要の順番）  .xlsx"
in_filename = "./temp/シラバス情報（修正版）.xlsx"
out_filename = "新旧対応表（情報）.xlsx"
in_filename2 = "./temp/シラバス情報（修正版）.xlsx"
out_filename2 = "新旧対応表（事業創造）.xlsx"

# Excelの中での読み書きのセル位置情報
# 授業科目の文字の入ったセルの位置を(1,1)とする
# 講義回数によって変化するから生成する

# out_mapの生成関数
def create_in_map(kougi_kaisu):
    map = [
        # 担当教員名〜時間数
        (1,3), (4,1), (3,5), (4,5), (5,5), (3,7), (4,7), (5, 7),
        # 概要，学習目標
        (7,1), (9,1),
    ]
    # 単元〜学習方法
    for i in range(kougi_kaisu) :
        map.append((12+i, 2))
        map.append((12+i, 7))
        map.append((12+i, 8))
    # 残りを追加
    map.extend( [
        # 使用図書〜発行年
        (28,3), (28,5), (28,7), (28,9),
        (29,3), (29,5), (29,7), (29,9),
        (30,3), (30,5), (30,7), (30,9),
        # 準備学修
        (31,3),
        # 評価方法・履修上の留意点
        (33,1), (33,4)
    ])
    return map

# out_mapの生成関数
def create_out_map(kougi_kaisu):
    map = [
        # 担当教員名〜時間数
        (1,2), (2,2), (3,2), (4,2), (5,2), (6,2), (7,2), (8,2),
        # 概要，学習目標
        (9,2), (10,2),
    ]
    # 単元〜学習方法
    for i in range(kougi_kaisu) :
        map.append((12+i, 2))
        map.append((12+i, 3))
        map.append((12+i, 4))
    # 残りを追加
    map.extend( [
        # 使用図書〜発行年
        (28,2), (28,3), (28,4), (28,5),
        (29,2), (29,3), (29,4), (29,5),
        (30,2), (30,3), (30,4), (30,5),
        # 準備学修
        (31,2),
        # 評価方法・履修上の留意点
        (32,2), (33,2)
        ])
    return map

# mapの生成
in_map_list = []
in_map_list.append(create_in_map(15))
out_map_list = []
out_map_list.append(create_out_map(15))

syllabus(in_filename, out_filename, in_map_list, out_map_list)
syllabus(in_filename2, out_filename2, in_map_list, out_map_list)


