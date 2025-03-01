import openpyxl as px
from my_logger import set_logger, getLogger
import time
import os
from dotenv import load_dotenv

#start = time.time()  # 現在時刻（処理開始前）を取得

load_dotenv()
set_logger()
logger = getLogger(__name__)

# 集計用Excelファイル
RESULT_EXCEL = os.getenv('RESULT_EXCEL')
EXCEL_SHEET = os.getenv('EXCEL_SHEET')

# 順位毎の桁表示
place_dict = {
    1 : 1000,
    2 :  100,
    3 :   10,
    4 :    1
}

# Excelファイルを読み込む
try:
    wb = px.load_workbook(RESULT_EXCEL)
    ws = wb[EXCEL_SHEET]
except FileNotFoundError:
    print(f"ファイル '{RESULT_EXCEL}' が見つかりません。")
    logger.error(f"ファイル '{RESULT_EXCEL}' が見つかりません。")
    exit()
except KeyError:
    print(f"シート '{EXCEL_SHEET}' が見つかりません。")
    logger.error(f"シート '{EXCEL_SHEET}' が見つかりません。")
    exit()

# 全データを取得
all_data = list(ws.iter_rows(min_row=2, min_col=1, max_col=7, values_only=True))
wb.close()

# 重複排除したプレイヤーリストを取得
name_list = sorted(set(row[3] for row in all_data))

# 結果格納リストの初期化
result_list = []

print()
print("==========通算成績==========")

# 順位点と順位を出力
def calculate_rank_points(all_data, name, place_dict):
    rank_points = 0
    place = 0
    for row in all_data:
        if str(name) == row[3]:
            rank_points = round(rank_points + row[5], 2)
            place += place_dict.get(row[2], 0)
    return rank_points, '【' + f'{place:04}' + '】'

# 一人ずつ処理
for name in name_list:
    rank_points, place = calculate_rank_points(all_data, name, place_dict)
    result_list.append([name, rank_points, place])

# 結果格納リストを順位点でソート
result_list = sorted(result_list, reverse=True, key=lambda x: x[1])  

# 1位から名前、順位点、順位を出力
count = 1
for x in result_list:
    name_padding = '\t' if len(x[0].encode('utf-8')) < 17 else '' # 名前が短い人はタブを1つ追加
    print(f"{count: >2}  {x[0]}{name_padding}\t{round(float(x[1]),2): >6}\t{x[2]}")
    count += 1

#end = time.time()  # 現在時刻（処理完了後）を取得
#time_diff = end - start  # 処理完了後の時刻から処理開始前の時刻を減算する
#print(time_diff)  # 処理にかかった時間データを使用