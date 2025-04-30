from datetime import datetime, timedelta
import pandas as pd

# 元のデータ（日付と時刻のログ）
raw_data = """
清水真二 2025/4/1 20:36:49
清水在宅作業
終了します。

清水真二 2025/4/2 09:51:23
清水在宅作業
開始します。

清水真二 2025/4/2 20:02:39
清水在宅作業
終了します。

清水真二 2025/4/3 09:00:35
清水在宅作業
開始します。

清水真二 2025/4/3 19:14:12
清水在宅作業
終了します。

清水真二 2025/4/4 10:26:18
清水在宅作業
開始します。

清水真二 2025/4/4 19:36:32
清水在宅作業
終了します。

清水真二 2025/4/7 09:06:21
清水在宅作業
開始します。

清水真二 2025/4/7 18:41:12
清水在宅作業
終了します。

清水真二 2025/4/8 10:52:02
清水在宅作業
開始します。

清水真二 2025/4/8 20:03:18
清水在宅作業
終了します。

清水真二 2025/4/9 10:10:07
清水在宅作業
開始します。

清水真二 2025/4/9 19:14:01
清水在宅作業
終了します。

清水真二 2025/4/10 10:21:30
清水在宅作業
開始します。

清水真二 2025/4/10 19:43:20
清水在宅作業
終了します。

清水真二 2025/4/11 09:02:06
清水在宅作業
開始します。

清水真二 2025/4/11 18:15:47
清水在宅作業
終了します。

清水真二 2025/4/14 09:07:31
清水在宅作業
開始します。

清水真二 2025/4/14 18:20:15
清水在宅作業
終了します。

清水真二 2025/4/15 10:02:55
清水在宅作業
開始します。

清水真二 2025/4/15 19:12:38
清水在宅作業
終了します。

清水真二 2025/4/16 09:23:55
清水在宅作業
開始します。

清水真二 2025/4/16 18:25:11
清水在宅作業
終了します。

清水真二 2025/4/17 09:50:53
清水在宅作業
開始します。

清水真二 2025/4/17 18:54:16
清水在宅作業
終了します。

清水真二 2025/4/18 09:53:20
清水在宅作業
開始します。

清水真二 2025/4/18 19:03:21
清水在宅作業
終了します。

清水真二 2025/4/21 09:01:34
清水在宅作業
開始します。

清水真二 2025/4/21 18:01:15
清水在宅作業
終了します。

清水真二 2025/4/22 09:55:02
清水在宅作業
開始します。

清水真二 2025/4/22 19:10:06
清水在宅作業
終了します。

清水真二 2025/4/23 09:52:20
清水在宅作業
開始します。

清水真二 2025/4/23 19:06:26
清水在宅作業
終了します。

清水真二 2025/4/24 09:26:19
清水在宅作業
開始します。

清水真二 2025/4/24 18:48:54
清水在宅作業
終了します。

清水真二 2025/4/25 09:30:12
清水在宅作業
開始します。

清水真二 2025/4/25 18:52:51
清水在宅作業
終了します。

清水真二 2025/4/28 09:02:29
清水在宅作業
開始します。

清水真二 2025/4/28 18:39:15
清水在宅作業
終了します。

清水真二 09:04:06
清水在宅作業
開始します。

清水真二 18:09:57
清水在宅作業
終了します。
"""




def main():
    print("Hello from time-card!")

    # 日時とメッセージのペアにパースする
    lines = raw_data.strip().split("\n")
    entries = []
    i = 0
    while i < len(lines):
        if "清水真二" in lines[i]:
            try:
                dt = lines[i].replace("清水真二", "").strip()
                if not dt.startswith("2025"):
                    dt = "2025/4/30 " + dt
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
    start_date = datetime(2025, 4, 1)
    date_rows = []
    for day_offset in range(30):
        current_date = start_date + timedelta(days=day_offset)
        weekday = ["月", "火", "水", "木", "金", "土", "日"][current_date.weekday()]
        date = current_date.day
        has_data = current_date.date() in log_dict
        row = {
            "日": date,
            "曜日": weekday,
            "休暇": "",
            "案件名": "清水在宅作業" if current_date.date() in log_dict else "",
            "開始時間": log_dict.get(current_date.date(), {}).get("開始", ""),
            "終了時間": log_dict.get(current_date.date(), {}).get("終了", ""),
            "昼休み": "1:00" if has_data else "",
        }
        date_rows.append(row)

    # データフレームに変換
    df = pd.DataFrame(date_rows)
    df.head(30)
    print(df.head(30))
    output_file = "output.xlsx"
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"Excelファイルを出力しました: {output_file}")


if __name__ == "__main__":
    main()
