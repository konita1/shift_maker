import pandas as pd
import calendar
from datetime import datetime, timedelta
import xlsxwriter
import argparse

parser = argparse.ArgumentParser()


parser.add_argument("--extra_name", type=str, default="社員S")
parser.add_argument("--extra_off_days", type=int, default=8)
parser.add_argument("--extra_start", type=str, default="06:00")
parser.add_argument("--extra_end", type=str, default="20:30")
parser.add_argument("--year", type=int, default=2026)
parser.add_argument("--month", type=int, default=5)

args = parser.parse_args()

extra_name = args.extra_name
extra_off_days = args.extra_off_days
extra_start = args.extra_start
extra_end = args.extra_end

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
# year = 2026
# month = 4
year = args.year
month = args.month

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
# 追加従業員ロジック
# 休日処理後の不足スロットを見て配置
# 条件：
# ・固定曜日なし
# ・月の休み日数を指定
# ・連続勤務は最大5日
# ・出勤日は8時間勤務＋1時間休憩
# ・1人以下の時間帯を優先
# =========================

extra_name = args.extra_name
extra_off_days = args.extra_off_days
extra_start = args.extra_start
extra_end = args.extra_end

extra_start_min = to_minutes(extra_start)
extra_end_min = to_minutes(extra_end)

days_in_month = calendar.monthrange(year, month)[1]
required_work_days = days_in_month - extra_off_days

WORK_DURATION = 8 * 60
BREAK_DURATION = 60
TOTAL_DURATION = 9 * 60


# =========================
# ① 連続勤務チェック
# =========================
def can_add_workday(selected_dates, candidate_date):

    temp_dates = selected_dates + [candidate_date]
    temp_dates = sorted([
        datetime.strptime(d, "%Y/%m/%d")
        for d in temp_dates
    ])

    consecutive = 1

    for i in range(1, len(temp_dates)):
        if (temp_dates[i] - temp_dates[i - 1]).days == 1:
            consecutive += 1
            if consecutive > 5:
                return False
        else:
            consecutive = 1

    return True


# =========================
# ② 日ごとの不足度
# =========================
all_dates = [
    datetime(year, month, day).strftime("%Y/%m/%d")
    for day in range(1, days_in_month + 1)
]

daily_scores = []

for date_str in all_dates:
    daily = schedule_df[schedule_df["date"] == date_str]

    score = 0

    for slot_start, slot_end, slot_name in slots:
        count = len(daily[daily["slot"] == slot_name])

        if count <= 1:
            score += 1

    daily_scores.append({
        "date": date_str,
        "score": score
    })

daily_scores = sorted(
    daily_scores,
    key=lambda x: x["score"],
    reverse=True
)


# =========================
# ③ 出勤日決定
# =========================
selected_work_dates = []

for item in daily_scores:
    if len(selected_work_dates) >= required_work_days:
        break

    if can_add_workday(selected_work_dates, item["date"]):
        selected_work_dates.append(item["date"])


# =========================
# ④ 勤務割り当て（8h + 1h休憩）
# =========================
extra_schedule = []

for date_str in selected_work_dates:

    daily = schedule_df[schedule_df["date"] == date_str]

    slot_scores = []

    for slot_start, slot_end, slot_name in slots:

        ss = to_minutes(slot_start)
        se = to_minutes(slot_end)

        if se <= extra_start_min or ss >= extra_end_min:
            continue

        count = len(daily[daily["slot"] == slot_name])

        score = max(0, 3 - count)

        slot_scores.append({
            "slot": slot_name,
            "start": ss,
            "score": score
        })

    if not slot_scores:
        continue

    best_slot = max(slot_scores, key=lambda x: x["score"])
    center_time = best_slot["start"]

    work_start = center_time - WORK_DURATION // 2
    work_start = (work_start // 30) * 30

    # =========================
    # ⑤ 勤務開始ルール
    # =========================
    six_am_min = to_minutes("06:00")
    normal_min_start = to_minutes("08:30")
    six_am_slot = "06:00-06:30"

    six_am_count = len(daily[daily["slot"] == six_am_slot])

    if six_am_count <= 1:
        work_start = six_am_min
    else:
        work_start = max(work_start, extra_start_min, normal_min_start)

    work_end = work_start + TOTAL_DURATION

    if work_end > extra_end_min:
        work_end = extra_end_min
        work_start = work_end - TOTAL_DURATION
        work_start = (work_start // 30) * 30

    # =========================
    # 休憩候補を作成
    # 基本：
    # ・00分出勤 → 勤務開始から4時間後
    # ・30分出勤 → 勤務開始から3.5時間後
    # ・探索範囲は勤務開始3時間後〜勤務終了3時間前
    # =========================
    # =========================
    # 既存休憩との重なり数をカウント
    # =========================
    # def count_existing_breaks(date_str, candidate_break_start, candidate_break_end):
    #     count = 0

    #     for item in break_conflicts:
    #         if item["date"] != date_str:
    #             continue

    #         slot_start, slot_end = item["slot"].split("-")
    #         bs = to_minutes(slot_start)
    #         be = to_minutes(slot_end)

    #         # 重なり判定
    #         if not (candidate_break_end <= bs or candidate_break_start >= be):
    #             count += 1

    #     return count
    # break_candidates = []

    # candidate_start = work_start + 3 * 60
    # candidate_limit = work_end - 3 * 60 - BREAK_DURATION

    # # 休憩の基本開始時刻
    # if work_start % 60 == 0:
    #     preferred_break_start = work_start + 4 * 60
    # else:
    #     preferred_break_start = work_start + 3 * 60 + 30

    # candidate = candidate_start

    # while candidate <= candidate_limit:
    #     candidate_end = candidate + BREAK_DURATION

    #     overlap_count = count_existing_breaks(
    #         date_str,
    #         candidate,
    #         candidate_end
    #     )

    #     # 基本休憩時刻から近いほど優先
    #     distance_from_preferred = abs(candidate - preferred_break_start)

    #     break_candidates.append({
    #         "break_start": candidate,
    #         "break_end": candidate_end,
    #         "overlap_count": overlap_count,
    #         "distance": distance_from_preferred
    #     })

    #     candidate += 30


    # best_break = min(
    #     break_candidates,
    #     key=lambda x: (x["overlap_count"], x["distance"])
    # )

    # break_start = best_break["break_start"]
    # break_end = best_break["break_end"]

    # =========================
    # 勤務開始時間と休憩時間を同時に決定
    # 休憩が被る場合は勤務開始を30分/60分ずらす
    # =========================

    def count_existing_breaks(date_str, candidate_break_start, candidate_break_end):
        count = 0

        for item in break_conflicts:
            if item["date"] != date_str:
                continue

            slot_start, slot_end = item["slot"].split("-")
            bs = to_minutes(slot_start)
            be = to_minutes(slot_end)

            if not (candidate_break_end <= bs or candidate_break_start >= be):
                count += 1

        return count


    def find_best_break_for_work_start(date_str, work_start, work_end):
        break_candidates = []

        candidate_start = work_start + 3 * 60
        candidate_limit = work_end - 3 * 60 - BREAK_DURATION

        if work_start % 60 == 0:
            preferred_break_start = work_start + 4 * 60
        else:
            preferred_break_start = work_start + 3 * 60 + 30

        candidate = candidate_start

        while candidate <= candidate_limit:
            candidate_end = candidate + BREAK_DURATION

            overlap_count = count_existing_breaks(
                date_str,
                candidate,
                candidate_end
            )

            distance_from_preferred = abs(candidate - preferred_break_start)

            break_candidates.append({
                "break_start": candidate,
                "break_end": candidate_end,
                "overlap_count": overlap_count,
                "distance": distance_from_preferred
            })

            candidate += 30

        if not break_candidates:
            return None

        return min(
            break_candidates,
            key=lambda x: (x["overlap_count"], x["distance"])
        )


    # まず基準の勤務開始時刻を決める
    base_work_start = work_start

    # 6:00開始の日は、基本的に6:00を優先
    # 通常日は、基準時刻・+30分・+60分を試す
    work_start_candidates = [
        base_work_start,
        base_work_start + 30,
        base_work_start + 60
    ]

    valid_candidates = []

    for candidate_work_start in work_start_candidates:

        candidate_work_end = candidate_work_start + TOTAL_DURATION

        # 勤務可能時間外なら除外
        if candidate_work_start < extra_start_min:
            continue

        if candidate_work_end > extra_end_min:
            continue

        best_break_candidate = find_best_break_for_work_start(
            date_str,
            candidate_work_start,
            candidate_work_end
        )

        if best_break_candidate is None:
            continue

        valid_candidates.append({
            "work_start": candidate_work_start,
            "work_end": candidate_work_end,
            "break_start": best_break_candidate["break_start"],
            "break_end": best_break_candidate["break_end"],
            "overlap_count": best_break_candidate["overlap_count"],
            "distance": best_break_candidate["distance"],
            "start_shift_amount": abs(candidate_work_start - base_work_start)
        })


    # 候補がない場合は、元の勤務開始で強行
    if not valid_candidates:
        work_end = work_start + TOTAL_DURATION
        break_start = work_start + 4 * 60
        break_end = break_start + BREAK_DURATION

    else:
        best_candidate = min(
            valid_candidates,
            key=lambda x: (
                x["overlap_count"],       # 休憩被りが少ない
                x["start_shift_amount"],  # 勤務開始のズレが少ない
                x["distance"]             # 基本休憩時刻に近い
            )
        )

        work_start = best_candidate["work_start"]
        work_end = best_candidate["work_end"]
        break_start = best_candidate["break_start"]
        break_end = best_candidate["break_end"]

    for slot_start, slot_end, slot_name in slots:

        ss = to_minutes(slot_start)
        se = to_minutes(slot_end)

        if se <= work_start or ss >= work_end:
            continue

        is_break = not (se <= break_start or ss >= break_end)

        extra_schedule.append({
            "date": date_str,
            "slot": slot_name,
            "name": extra_name,
            "is_break": is_break
        })


# =========================
# ⑥ schedule_dfへ反映
# =========================
extra_schedule_df = pd.DataFrame(extra_schedule)

if not extra_schedule_df.empty:

    # 勤務だけ追加
    extra_work_df = extra_schedule_df[
        extra_schedule_df["is_break"] == False
    ][["date", "slot", "name"]]

    schedule_df = pd.concat(
        [schedule_df, extra_work_df],
        ignore_index=True
    )

    # 休憩は赤表示用
    extra_break_df = extra_schedule_df[
        extra_schedule_df["is_break"] == True
    ]

    for _, row in extra_break_df.iterrows():
        break_conflicts.append({
            "date": row["date"],
            "name": row["name"],
            "slot": row["slot"]
        })


# =========================
# ⑦ 列整理（超重要）
# =========================
schedule_df = schedule_df[["date", "slot", "name"]]
schedule_df = schedule_df.dropna(subset=["date", "slot", "name"])


# =========================
# ⑥ 念のため schedule_df の列を整理
# =========================
schedule_df = schedule_df[["date", "slot", "name"]]
schedule_df = schedule_df.dropna(subset=["date", "slot", "name"])

# print("追加従業員の出勤日:", selected_work_dates)
# print("追加従業員スロット数:", len(extra_schedule_df))
# print(extra_schedule_df.head())
# =========================
# ⑭ Excel出力（休憩赤表示付き）
# =========================
month_str = f"{year}_{month:02d}"
output_excel_path = f"shift/{month_str}_shift.xlsx"
# output_excel_path = "shift_by_day.xlsx"

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


    # all_employee_names = schedule_df["name"].unique()
    # all_employee_names = contracts["name"].unique()
    contract_employee_names = contracts["name"].drop_duplicates().tolist()

    all_employee_names = contract_employee_names.copy()

    if extra_name not in all_employee_names:
        all_employee_names.append(extra_name)
        
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

