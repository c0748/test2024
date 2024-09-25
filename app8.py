import openpyxl
import tkinter as tk
from tkinter import Label, Entry, Button
from PIL import Image, ImageTk
import requests
from io import BytesIO
from datetime import datetime, timedelta


def show_cycle_info():
    name = name_entry.get()

    standard_time = int(standard_entry.get())
    # エクセルを読み込む
    wb = openpyxl.load_workbook("files/colored_sequence_PLOT3.xlsx")
    ws = wb.active
    # ユーザーから日付を入力してもらう
    input_date = name
    # 現在の時間を取得する 午前０時からの経過時間 分表示
    now = datetime.now()
    print(now)
    midnight = datetime.combine(now.date(), datetime.min.time())
    # 経過時間を計算する
    elapsed_minutes = now - midnight

    # 停止時間は画面から入力
    stop_minutes = int(time_entry.get())
    # 稼働時間を計算
    # input_timeをtimedelta（分）として扱い、経過時間から引く
    adjusted_time = elapsed_minutes - timedelta(minutes=stop_minutes)
    print(adjusted_time)
    # 経過時間を秒に換算する
    elapsed_seconds = adjusted_time.total_seconds()
    # 稼働時間を standard_entry で割る
    result = elapsed_seconds / standard_time

    # #1 入力された日付のデータを抽出
    date_filter = []

    # 行数を取得する
    count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[4] == input_date:
            date_filter.append(row)
            count += 1

    if date_filter:
        # サイクルタイムを秒に変換してリストに格納
        cycle_times = [time_to_seconds(row[3]) for row in date_filter]
        # #3 ４列目の最大値を求める
        max_cycle_time = max(cycle_times)
        # #4 ４列目の最小値を求める
        min_cycle_time = min(cycle_times)
        # #5 ４列目の平均値を求める
        average_cycle_time = sum(cycle_times) / len(cycle_times)

        # サイクル数、最大、最小、平均を表示
        count_label.config(text=f"サイクル数: {count}")
        max_label.config(text=f"最大サイクルタイム:{max_cycle_time // 60}分{max_cycle_time % 60}秒")
        min_label.config(text=f"最小サイクルタイム: {min_cycle_time // 60}分{min_cycle_time % 60}秒")
        ave_label.config(
            text=f"平均サイクルタイム:{average_cycle_time // 60:.0f}分{average_cycle_time % 60:.0f}秒"
        )
        goal_label.config(text=f"目標サイクル数は：{int(result)}回")
    else:
        count_label.config(text=f"日付 '{count}' は見つかりませんでした。")
        max_label.config(text="")
        min_label.config(text="")
        ave_label.config(text="")


# 秒を〇分〇秒にするための計算
def time_to_seconds(time_str):
    minutes, seconds = map(int, time_str.split(":"))
    return minutes * 60 + seconds


# GUIの初期設定

root = tk.Tk()
root.title("サイクル情報表示")


# 日付入力フィールド
Label(root, text="日付を入力してください:", font=("Arial", 16)).pack(padx=20, pady=10)
name_entry = Entry(root, font=("Arial", 14))
name_entry.pack(padx=20, pady=10)

Label(root, text="停止時間を入力してください（分）:", font=("Arial", 16)).pack(padx=20, pady=10)
time_entry = Entry(root, font=("Arial", 14))
time_entry.pack(padx=20, pady=10)


Label(root, text="基準サイクルタイムを入力してください（秒）:", font=("Arial", 16)).pack(padx=20, pady=10)
standard_entry = Entry(root, font=("Arial", 14))
standard_entry.pack(padx=20, pady=10)

# 表示ボタン

Button(root, text="表示", font=("Arial", 14), command=show_cycle_info).pack(padx=20, pady=10)


# サイクルタイム表示用ラベル（文字なし（中身なし））

count_label = Label(root, font=("Arial", 20))
count_label.pack(padx=20, pady=10)
max_label = Label(root, font=("Arial", 16))
max_label.pack(padx=20, pady=5)
min_label = Label(root, font=("Arial", 16))
min_label.pack(padx=20, pady=5)
ave_label = Label(root, font=("Arial", 16))
ave_label.pack(padx=20, pady=5)
goal_label = Label(root, font=("Arial", 20))
goal_label.pack(padx=20, pady=10)


# GUIの表示
root.mainloop()
