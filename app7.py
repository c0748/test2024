import openpyxl

# Excelファイルを読み込む
wb = openpyxl.load_workbook("files/colored_sequence_PLOT3.xlsx")
ws = wb.active

# ユーザーから日付を入力してもらう
input_date = input("日付を入力してください (例: 2024/08/31): ")
print(type(input_date))

# #1 入力された日付のデータを抽出
date_filter = []

# 行数を取得する
count = 0
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[4] == input_date:
        date_filter.append(row)
        count += 1


print(date_filter)

# サイクルタイムを秒に変換する関数
def time_to_seconds(time_str):
    minutes, seconds = map(int, time_str.split(":"))
    return minutes * 60 + seconds


# サイクルタイムを秒に変換してリストに格納
cycle_times = [time_to_seconds(row[3]) for row in date_filter]

# #3 ４列目の最大値を求める
max_cycle_time = max(cycle_times)

# #4 ４列目の最小値を求める
min_cycle_time = min(cycle_times)

# #5 ４列目の平均値を求める
average_cycle_time = sum(cycle_times) / len(cycle_times)

# 結果を表示
print(f"#1 入力された日付のデータ:\n{input_date}")
print(f"#2 行数: {count}")
print(f"#3 最大サイクルタイム: {max_cycle_time // 60}分{max_cycle_time % 60}秒")
print(f"#4 最小サイクルタイム: {min_cycle_time // 60}分{min_cycle_time % 60}秒")
print(f"#5 平均サイクルタイム: {average_cycle_time // 60:.0f}分{average_cycle_time % 60:.0f}秒")
