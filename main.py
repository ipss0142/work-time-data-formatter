from datetime import datetime, timedelta
import pandas as pd
import json


def main():
    print("Hello from time-card!")

    # 設定ファイルを読み込む
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)

    user_name = config["user_name"]
    project_name = config["project_name"]
    start_date_str = config["start_date"]
    output_file = config["output_file"]
    lunch_break = config["lunch_break"]

    # 勤怠データを読み込む
    with open("timecard.txt", "r", encoding="utf-8") as f:
        raw_data = f.read()

    # 日時とメッセージのペアにパースする
    lines = raw_data.strip().split("\n")
    entries = []
    i = 0
    while i < len(lines):
        if user_name in lines[i]:
            try:
                dt = lines[i].replace(user_name, "").strip()
                if not dt.startswith("2025"):
                    dt = "2025/6/30 " + dt
                timestamp = datetime.strptime(dt, "%Y/%m/%d %H:%M:%S")
                status = lines[i+2].strip()
                entries.append((timestamp, status))
            except Exception:
                pass
            i += 3
        else:
            i += 1

    # 日付ごとに開始・終了をマッピング
    log_dict = {}
    for ts, status in entries:
        day = ts.date()
        if day not in log_dict:
            log_dict[day] = {"開始": None, "終了": None}
        if status.startswith("開始"):
            log_dict[day]["開始"] = ts.strftime("%H:%M")
        elif status.startswith("終了"):
            log_dict[day]["終了"] = ts.strftime("%H:%M")

    # Excel形式に整形
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    date_rows = []
    for day_offset in range(31):
        current_date = start_date + timedelta(days=day_offset)
        weekday = ["月", "火", "水", "木", "金", "土", "日"][current_date.weekday()]
        date = current_date.day
        has_data = current_date.date() in log_dict
        row = {
            "日": date,
            "曜日": weekday,
            "休暇": "",
            "案件名": project_name if current_date.date() in log_dict else "",
            "開始時間": log_dict.get(current_date.date(), {}).get("開始", ""),
            "終了時間": log_dict.get(current_date.date(), {}).get("終了", ""),
            "昼休み": lunch_break if has_data else "",
        }
        date_rows.append(row)

    # データフレームに変換
    df = pd.DataFrame(date_rows)
    df.head(31)
    print(df.head(31))
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"Excelファイルを出力しました: {output_file}")


if __name__ == "__main__":
    main()
