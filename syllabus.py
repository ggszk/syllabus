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

# 入出力ファイル
in_filename = "syl_org.xlsx"
out_filename = "syl_out.xlsx"

# Excelの中での読み書きのセル位置情報
# 授業科目の文字の入ったセルの位置を(1,1)とする
in_map = [
    # 担当教員名〜時間数
    (1,3), (4,1), (3,5), (4,5), (5,5), (3,7), (4,7), (5, 7),
    # 概要，学習目標
    (7,1), (9,1),
    # 単元〜学習方法
    (12,2), (12,7), (12,8),
    (13,2), (13,7), (13,8),
    (14,2), (14,7), (14,8),
    (15,2), (15,7), (15,8),
    (16,2), (16,7), (16,8),
    (17,2), (17,7), (17,8),
    (18,2), (18,7), (18,8),
    (19,2), (19,7), (19,8),
    (20,2), (20,7), (20,8),
    (21,2), (21,7), (21,8),
    (22,2), (22,7), (22,8),
    (23,2), (23,7), (23,8),
    (24,2), (24,7), (24,8),
    (25,2), (25,7), (25,8),
    (26,2), (26,7), (26,8),
    # 使用図書〜発行年
    (28,3), (28,5), (28,7), (28,9),
    (29,3), (29,5), (29,7), (29,9),
    (30,3), (30,5), (30,7), (30,9),
    (31,3), (31,5), (31,7), (31,9),
    # 準備学修
    (32,3),
    # 評価方法・履修上の留意点
    (34,1), (34,4)
    ]
out_map = [
    # 担当教員名〜時間数
    (1,2), (2,2), (3,2), (4,2), (5,2), (6,2), (7,2), (8,2),
    # 概要，学習目標
    (9,2), (10,2),
    # 単元〜学習方法
    (12,2), (12,3), (12,4),
    (13,2), (13,3), (13,4),
    (14,2), (14,3), (14,4),
    (15,2), (15,3), (15,4),
    (16,2), (16,3), (16,4),
    (17,2), (17,3), (17,4),
    (18,2), (18,3), (18,4),
    (19,2), (19,3), (19,4),
    (20,2), (20,3), (20,4),
    (21,2), (21,3), (21,4),
    (22,2), (22,3), (22,4),
    (23,2), (23,3), (23,4),
    (24,2), (24,3), (24,4),
    (25,2), (25,3), (25,4),
    (26,2), (26,3), (26,4),
    # 使用図書〜発行年
    (28,2), (28,3), (28,4), (28,5),
    (29,2), (29,3), (29,4), (29,5),
    (30,2), (30,3), (30,4), (30,5),
    (31,2), (31,3), (31,4), (31,5),
    # 準備学修
    (32,2),
    # 評価方法・履修上の留意点
    (33,2), (34,2)
    ]

# ワークブック
in_wb = openpyxl.load_workbook(in_filename)
out_wb = openpyxl.load_workbook(out_filename)

# 科目
#kamoku_list = ['確率論', 'データベースの基礎']
kamoku_list = ['確率論']

# セルの読み書き
for kamoku in kamoku_list :
    # シート
    in_sheet = in_wb[kamoku]
    out_sheet = out_wb[kamoku]

    # オフセットは，「授業科目」が左上(1,1)の位置からいくつずれているか（ずれなしは(0,0)）
    # セルを読む：入力
    in_cells = read_cells(in_sheet, in_map, (1,0))

    # セルに代入
    # 新
    write_cells(out_sheet, in_cells, out_map, (2,1))
    # 旧
    write_cells(out_sheet, in_cells, out_map, (2,8))

# ファイル保存
out_wb.save(out_filename)
