import pandas as pd
import calendar
from datetime import datetime, timedelta
import xlsxwriter

# ② データ読み込み
contracts = pd.read_csv("data/contracts.csv")
holidays = pd.read_csv("data/holidays.csv")

# print(contracts.head())

# =========================
# ③ 休み整形
# =========================
holidays["date"] = pd.to_datetime(holidays["date"]).dt.strftime("%Y/%m/%d")


# =========================
# ④ カレンダー生成
# =========================
year = 2026
month = 4

cal = calendar.monthcalendar(year, month)
df_cal = pd.DataFrame(cal, columns=["Mon","Tue","Wed","Thu","Fri","Sat","Sun"])


# =========================
# ⑤ 30分スロット生成
# =========================
def generate_slots(start="06:00", end="20:30"):
    slots = []
    current = datetime.strptime(start, "%H:%M")
    end_dt = datetime.strptime(end, "%H:%M")

    while current < end_dt:
        nxt = current + timedelta(minutes=30)
        slots.append((
            current.strftime("%H:%M"),
            nxt.strftime("%H:%M"),
            current.strftime("%H:%M") + "-" + nxt.strftime("%H:%M")
        ))
        current = nxt
    return slots

slots = generate_slots()


# =========================
# ⑥ 時間変換
# =========================
def to_minutes(t):
    if pd.isna(t):
        return None  # または 0 や適切な値

    t = str(t)

    # "600" → "6:00" に補正したい場合（必要なら）
    if ":" not in t:
        t = t.zfill(4)
        t = t[:2] + ":" + t[2:]

    h, m = map(int, t.split(":"))
    return h * 60 + m


# =========================
# ⑦ 休憩判定
# =========================
def is_in_break(slot_start, slot_end, break_start, break_end):
    if pd.isna(break_start):
        return False

    ss = to_minutes(slot_start)
    se = to_minutes(slot_end)
    bs = to_minutes(break_start)
    be = to_minutes(break_end)

    return not (se <= bs or ss >= be)


# =========================
# ⑧ 休憩記録用
# =========================
break_conflicts = []


# =========================
# ⑨ 勤務可能判定
# =========================
def is_available(contract_row, slot_start, slot_end, date_str):

    cs = contract_row["start_min"]
    ce = contract_row["end_min"]

    ss = to_minutes(slot_start)
    se = to_minutes(slot_end)

    # 勤務時間外
    if ce <= ss or cs >= se:
        return False

    # 休憩チェック
    if is_in_break(
        slot_start,
        slot_end,
        contract_row.get("break_start"),
        contract_row.get("break_end")
    ):
        break_conflicts.append({
            "date": date_str,
            "name": contract_row["name"],
            "slot": slot_start + "-" + slot_end
        })
        return False

    return True


# =========================
# ⑩ 契約前処理
# =========================
contracts["start_min"] = contracts["start"].apply(to_minutes)
contracts["end_min"] = contracts["end"].apply(to_minutes)


# =========================
# ⑪ 日付変換
# =========================
def to_date(day):
    if day == 0:
        return None
    return datetime(year, month, day).strftime("%Y/%m/%d")


# =========================
# ⑫ シフト生成
# =========================
schedule = []
# break_conflictsリストをここでクリアして、毎回新しいシフトデータのみを反映させる
break_conflicts = []

for i in range(len(df_cal)):
    for col in df_cal.columns:

        day = df_cal.loc[i, col]
        date_str = to_date(day)

        if date_str is None:
            continue

        weekday = col

        for _, emp in contracts.iterrows():

            if emp["weekday"] != weekday:
                continue

            for slot_start, slot_end, slot_name in slots:

                if is_available(emp, slot_start, slot_end, date_str):

                    schedule.append({
                        "date": date_str,
                        "slot": slot_name,
                        "name": emp["name"]
                    })


schedule_df = pd.DataFrame(schedule)


# =========================
# ⑬ 休み反映
# =========================
schedule_df = schedule_df.merge(
    holidays,
    on=["name", "date"],
    how="left",
    indicator=True
)

schedule_df = schedule_df[schedule_df["_merge"] == "left_only"]
schedule_df = schedule_df.drop(columns=["_merge"])


# =========================
# 休憩データも休日除外
# =========================
break_df = pd.DataFrame(break_conflicts)

if not break_df.empty:
    break_df = break_df.merge(
        holidays,
        on=["name", "date"],
        how="left",
        indicator=True
    )

    break_df = break_df[break_df["_merge"] == "left_only"]
    break_df = break_df.drop(columns=["_merge"])

    break_conflicts = break_df.to_dict("records")

# =========================
# ⑭ Excel出力（休憩赤表示付き）
# =========================
output_excel_path = "shift_by_day.xlsx"

with pd.ExcelWriter(output_excel_path, engine="xlsxwriter") as writer:

    workbook = writer.book

    red_format = workbook.add_format({
        "bg_color": "#FFC7CE",
        "font_color": "#FFC7CE"
    })

    all_slot_names = pd.DataFrame(
        slots,
        columns=["start", "end", "slot"]
    )["slot"]

    all_employee_names = contracts["name"].unique()

    for date_val in schedule_df["date"].unique():

        daily = schedule_df[schedule_df["date"] == date_val]

        pivot = daily.pivot_table(
            index="name",
            columns="slot",
            values="name",
            aggfunc=lambda x: "□",
            fill_value=""
        )

        pivot = pivot.reindex(
            index=all_employee_names,
            columns=all_slot_names,
            fill_value=""
        )

        sheet_name = str(date_val).replace("/", "-")
        pivot.to_excel(writer, sheet_name=sheet_name)

        worksheet = writer.sheets[sheet_name]

        # 現在のシート（日付）に該当する休憩衝突のみをフィルタリング
        current_date_break_conflicts = [
            item for item in break_conflicts if item["date"] == date_val
        ]

        # 休憩セルを赤表示
        for r, row in enumerate(pivot.index, start=1):
            for c, col in enumerate(pivot.columns, start=1):
                for item in current_date_break_conflicts:
                    if item["name"] == row and item["slot"] == col:
                        worksheet.write(r, c, "■", red_format)

print(f"Excel出力完了：{output_excel_path}")