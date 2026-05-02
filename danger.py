import pandas as pd
import argparse
from datetime import datetime, timedelta

parser = argparse.ArgumentParser()

parser.add_argument("--year", type=int, default=2026)
parser.add_argument("--month", type=int, default=5)
args = parser.parse_args()

year = args.year
month = args.month

month_str = f"{year}_{month:02d}"
file_path = f"shift/{month_str}_shift.xlsx"


xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

def generate_slots(start="06:00", end="20:30"):
    slots = []
    current = datetime.strptime(start, "%H:%M")
    end_dt = datetime.strptime(end, "%H:%M")

    while current < end_dt:
        nxt = current + timedelta(minutes=30)
        slots.append(
            current.strftime("%H:%M") + "-" + nxt.strftime("%H:%M")
        )
        current = nxt

    return slots

all_slot_names = generate_slots()

# =========================
# ① 結果格納
# =========================
danger_slots = []   # 1人以下
warning_slots = []  # ちょうど2人


# =========================
# ② チェック処理
# =========================
# for sheet in sheet_names:

#     df = pd.read_excel(file_path, sheet_name=sheet, index_col=0)

#     for col in df.columns:

#         count = (df[col] == "□").sum()

#         # 1人以下（危険）
#         if count <= 1:
#             danger_slots.append({
#                 "date": sheet,
#                 "slot": col,
#                 "staff_count": count
#             })

#         # ちょうど2人（注意）
#         elif count == 2:
#             warning_slots.append({
#                 "date": sheet,
#                 "slot": col,
#                 "staff_count": count
#             })
for sheet in sheet_names:

    # 月間シフト表は除外
    if sheet == "月間シフト表":
        continue

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    # 0行目: 日付・曜日
    # 1行目: 氏名・時間
    # 2行目以降: 従業員
    employee_df = df.iloc[2:, :]

    # 下部集計行を除外
    employee_df = employee_df[
        ~employee_df.iloc[:, 0].isin([
            "社員S時間",
            "PT時間",
            "合計時間",
            "追加入時間",
            "総人数"
        ])
    ]

    # スロット列だけ取得
    slot_df = employee_df.iloc[:, 1:1 + len(all_slot_names)]

    for i, slot_name in enumerate(all_slot_names):

        count = (slot_df.iloc[:, i] == "□").sum()

        if count <= 1:
            danger_slots.append({
                "date": sheet,
                "slot": slot_name,
                "staff_count": count
            })

        elif count == 2:
            warning_slots.append({
                "date": sheet,
                "slot": slot_name,
                "staff_count": count
            })


# ========================

# =========================
# ④ 出力②：注意（2人）
# =========================
warning_df = pd.DataFrame(warning_slots)

print("🟡【注意】2人のスロット")
if warning_df.empty:
    print("なし（問題ありません）")
else:
    warning_df = warning_df.sort_values(["date", "slot"])
    print(warning_df)

# ③ 出力①：危険（1人以下）
# =========================
danger_df = pd.DataFrame(danger_slots)

print("🔴【危険】1人以下のスロット")
if danger_df.empty:
    print("なし（問題ありません）")
else:
    danger_df = danger_df.sort_values(["date", "slot"])
    print(danger_df)



# file_path = "shift_by_day.xlsx"
# output_path = "danger_days_shift.xlsx"
file_path = f"shift/{month_str}_shift.xlsx"
output_path = f"shift/{month_str}_shift_danger.xlsx"

xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names


# # =========================
# # ① 危険日抽出
# # =========================
# danger_dates = []

# for sheet in sheet_names:

#     df = pd.read_excel(file_path, sheet_name=sheet, index_col=0)

#     for col in df.columns:
#         if (df[col] == "□").sum() <= 1:
#             danger_dates.append(sheet)
#             break

# danger_dates = set(danger_dates)
# =========================
# ① 危険日抽出
# =========================
danger_dates = []

for sheet in sheet_names:

    # 月間シフト表・月シートは除外
    if sheet in ["月間シフト表", "月"]:
        continue

    df = pd.read_excel(file_path, sheet_name=sheet, header=None)

    employee_df = df.iloc[2:, :]

    employee_df = employee_df[
        ~employee_df.iloc[:, 0].isin([
            "社員S時間",
            "PT時間",
            "合計時間",
            "追加入時間",
            "総人数"
        ])
    ]

    slot_df = employee_df.iloc[:, 1:1 + len(all_slot_names)]

    has_danger = False

    for i in range(len(all_slot_names)):
        count = (slot_df.iloc[:, i] == "□").sum()

        if count <= 1:
            has_danger = True
            break

    if has_danger:
        danger_dates.append(sheet)

danger_dates = set(danger_dates)

# =========================
# ② Excel出力（再描画＋色付け）
# =========================
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:

    workbook = writer.book

    # 休憩・リスク共通フォーマット
    red_format = workbook.add_format({
        "bg_color": "#FFC7CE",
        "font_color": "#FFC7CE"
    })
    red_format_time = workbook.add_format({
        "bg_color": "#FFC7CE",
        "font_color": "#9C0006"
    })

    # for sheet in sheet_names:

    #     # 危険日だけ出力
    #     if sheet not in danger_dates:
    #         continue

    #     df = pd.read_excel(file_path, sheet_name=sheet, index_col=0)
    #     df.to_excel(writer, sheet_name=sheet)

    for sheet in sheet_names:

        if sheet not in danger_dates:
            continue

        df = pd.read_excel(file_path, sheet_name=sheet, header=None)
        df.to_excel(writer, sheet_name=sheet, index=False, header=False)

        worksheet = writer.sheets[sheet]

        # =========================
        # ③ 休憩（■）を赤
        # =========================
        for r, row in enumerate(df.index, start=1):
            for c, col in enumerate(df.columns, start=1):

                if df.loc[row, col] == "■":
                    worksheet.write(r, c, "■", red_format)

        # =========================
        # ④ 人数不足スロット（最上段のみ赤）
        # =========================
        for c, col in enumerate(df.columns, start=1):

            count = (df[col] == "□").sum()

            if count <= 1:
                worksheet.write(0, c, col, red_format_time)


print(f"不足日のみ（色付き）出力完了：{output_path}")