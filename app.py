
import io
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st


# =========================
# 基本設定
# =========================
st.set_page_config(page_title="製作時程排程工具", page_icon="📅", layout="wide")

DEFAULT_PROJECT_NAME = "Dove多芬_十效修護精華髮油Q2宣傳"
DEFAULT_MODE = "從開始日期往後排"
DEFAULT_START_DATE = date.today()
DEFAULT_LAUNCH_DATE = date(2026, 5, 6)
DEFAULT_COLLAPSE_THRESHOLD = 2

DEFAULT_TASKS = [
    {"任務名稱": "提供素材", "負責人": "客戶", "工作天數": 1, "上線日": False},
    {"任務名稱": "視覺製作", "負責人": "Ad2", "工作天數": 3, "上線日": False},
    {"任務名稱": "客戶確認", "負責人": "客戶", "工作天數": 1, "上線日": False},
    {"任務名稱": "視覺調整", "負責人": "Ad2", "工作天數": 2, "上線日": False},
    {"任務名稱": "客戶確認", "負責人": "客戶", "工作天數": 1, "上線日": False},
    {"任務名稱": "廣告進稿", "負責人": "Ad2", "工作天數": 1, "上線日": False},
    {"任務名稱": "廣告上線", "負責人": "Ad2", "工作天數": 1, "上線日": True},
]

DEFAULT_HOLIDAYS = {
    '2025-12-25': '行憲紀念日',
    '2026-01-01': '元旦',
    '2026-01-21': 'Ad2尾牙',
    '2026-02-15': '春節連假',
    '2026-02-16': '春節連假',
    '2026-02-17': '春節連假',
    '2026-02-18': '春節連假',
    '2026-02-19': '春節連假',
    '2026-02-20': '春節連假',
    '2026-02-27': '二二八連假',
    '2026-02-28': '二二八連假',
    '2026-04-03': '清明連假',
    '2026-04-04': '清明連假',
    '2026-04-05': '清明連假',
    '2026-04-06': '清明連假',
    '2026-05-01': '勞動節放假',
    '2026-06-19': '端午節連假',
    '2026-09-25': '國定連假',
    '2026-09-28': '國定連假',
    '2026-10-09': '國慶日連假',
    '2026-10-10': '國慶日連假',
    '2026-10-25': '光復節連假',
    '2026-10-26': '光復節連假',
    '2026-12-25': '行憲紀念日連假',
}

MODE_MAP = {
    "從開始日期往後排": "forward",
    "從上線日期往前推": "backward",
    "同時指定開始與上線日期": "double",
}

COLOR_CLIENT_BAR = '#EA9B56'
COLOR_AD2_BAR = '#4BACC6'
COLOR_LAUNCH_BAR = '#FF0000'
COLOR_PREP_BAR = '#92D050'
COLOR_WEEKEND = '#D9D9D9'
COLOR_LEGEND_CLIENT = '#EA9B56'
COLOR_LEGEND_AD2 = '#4BACC6'
COLOR_BREAK_TEXT = '#000000'
COLOR_HOLIDAY_TEXT = '#595959'
COLOR_HOLIDAY_BG = COLOR_WEEKEND
MONTH_COLORS = ['#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#E7E6E6']


# =========================
# Session state
# =========================
if "tasks_df" not in st.session_state:
    st.session_state.tasks_df = pd.DataFrame(DEFAULT_TASKS)

if "holidays_text" not in st.session_state:
    st.session_state.holidays_text = "\n".join(
        [f"{k},{v}" for k, v in DEFAULT_HOLIDAYS.items()]
    )

if "schedule_df" not in st.session_state:
    st.session_state.schedule_df = None

if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

if "warning_msg" not in st.session_state:
    st.session_state.warning_msg = ""


# =========================
# 工具函式
# =========================
def parse_holidays(text: str) -> dict:
    holidays = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if "," not in line:
            raise ValueError(f"假日格式錯誤：{line}，請使用 YYYY-MM-DD,名稱")
        d, name = line.split(",", 1)
        d = d.strip()
        name = name.strip()
        pd.to_datetime(d)  # 驗證格式
        holidays[d] = name
    return holidays


def normalize_tasks(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        raise ValueError("請至少保留一筆任務。")

    df = df.copy()

    required_cols = ["任務名稱", "負責人", "工作天數", "上線日"]
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"缺少欄位：{c}")

    df["任務名稱"] = df["任務名稱"].fillna("").astype(str).str.strip()
    df["負責人"] = df["負責人"].fillna("Ad2").astype(str).str.strip()
    df["工作天數"] = pd.to_numeric(df["工作天數"], errors="coerce")
    df["上線日"] = df["上線日"].fillna(False).astype(bool)

    df = df[df["任務名稱"] != ""].copy()
    if df.empty:
        raise ValueError("請至少保留一筆有任務名稱的資料。")

    if (df["工作天數"].isna()).any() or (df["工作天數"] <= 0).any():
        raise ValueError("工作天數必須為大於 0 的整數。")

    launch_count = int(df["上線日"].sum())
    if launch_count > 1:
        raise ValueError("「上線日」只能勾選一筆。")
    if launch_count == 0:
        # 若未勾選，預設最後一筆視為上線日
        df.loc[df.index[-1], "上線日"] = True

    tasks = []
    for _, row in df.iterrows():
        tasks.append(
            {
                "task": row["任務名稱"],
                "owner": row["負責人"],
                "days": int(row["工作天數"]),
                "is_launch": bool(row["上線日"]),
            }
        )
    return tasks


def build_scheduler(tasks_config, holidays_config, calculation_mode, start_date_obj, launch_date_obj):
    holidays_dt = [pd.to_datetime(h).date() for h in holidays_config.keys()]

    def is_workday(d):
        return (d.weekday() < 5) and (d not in holidays_dt)

    def subtract_workdays(start_date, days):
        current = start_date
        if days == 0:
            return current
        check_days = days - 1
        while check_days > 0:
            current -= timedelta(days=1)
            if is_workday(current):
                check_days -= 1
        return current

    def add_workdays(start_date, days):
        current = start_date
        if days == 0:
            return current
        check_days = days - 1
        while check_days > 0:
            current += timedelta(days=1)
            if is_workday(current):
                check_days -= 1
        return current

    def get_previous_workday(d):
        d -= timedelta(days=1)
        while not is_workday(d):
            d -= timedelta(days=1)
        return d

    def get_next_workday(d):
        d += timedelta(days=1)
        while not is_workday(d):
            d += timedelta(days=1)
        return d

    def ensure_workday_forward(d):
        while not is_workday(d):
            d += timedelta(days=1)
        return d

    schedule = []
    warning_msg = ""

    if not tasks_config:
        raise ValueError("目前沒有可排程的任務。")

    launch_tasks = [t for t in tasks_config if t["is_launch"]]
    launch_task_config = launch_tasks[0] if launch_tasks else tasks_config[-1]

    if calculation_mode == "backward":
        if not launch_date_obj:
            raise ValueError("「從上線日期往前推」需要填寫上線日期。")

        current_end = launch_date_obj
        reversed_tasks = tasks_config[::-1]
        temp_schedule = []

        for i, t in enumerate(reversed_tasks):
            duration = t["days"]
            if t["is_launch"]:
                end_date = current_end
                start_date = subtract_workdays(end_date, duration)
            elif i == 0:
                end_date = current_end
                start_date = subtract_workdays(end_date, duration)
            else:
                prev_start = temp_schedule[-1]["Start Date"]
                end_date = get_previous_workday(prev_start)
                start_date = subtract_workdays(end_date, duration)

            temp_schedule.append(
                {
                    "Task": t["task"],
                    "Owner": t["owner"],
                    "Start Date": start_date,
                    "End Date": end_date,
                    "Type": "Launch" if t["is_launch"] else "Normal",
                }
            )

        schedule = temp_schedule[::-1]

    elif calculation_mode == "forward":
        curr_start = ensure_workday_forward(start_date_obj or date.today())
        prev_end = None

        for idx, t in enumerate(tasks_config):
            if t["is_launch"] and launch_date_obj:
                start_d = launch_date_obj
            else:
                start_d = curr_start if idx == 0 else get_next_workday(prev_end)

            end_d = add_workdays(start_d, t["days"])
            schedule.append(
                {
                    "Task": t["task"],
                    "Owner": t["owner"],
                    "Start Date": start_d,
                    "End Date": end_d,
                    "Type": "Launch" if t["is_launch"] else "Normal",
                }
            )
            prev_end = end_d

        if launch_date_obj:
            launch_item = next((x for x in schedule if x["Type"] == "Launch"), None)
            if launch_item and launch_item["Start Date"] != launch_date_obj:
                warning_msg = "⚠️ 上線日任務未能固定在指定日期，請檢查流程設定。"

    else:  # double
        if not start_date_obj or not launch_date_obj:
            raise ValueError("「同時指定開始與上線日期」需要同時填寫開始日期與上線日期。")

        curr_start = ensure_workday_forward(start_date_obj)
        prev_end = None

        normal_tasks = [t for t in tasks_config if not t["is_launch"]]

        for idx, t in enumerate(normal_tasks):
            start_d = curr_start if idx == 0 else get_next_workday(prev_end)
            end_d = add_workdays(start_d, t["days"])
            schedule.append(
                {
                    "Task": t["task"],
                    "Owner": t["owner"],
                    "Start Date": start_d,
                    "End Date": end_d,
                    "Type": "Normal",
                }
            )
            prev_end = end_d

        if schedule:
            last_task_end = schedule[-1]["End Date"]
        else:
            last_task_end = start_date_obj

        real_prep_start = last_task_end + timedelta(days=1)
        real_prep_end = launch_date_obj - timedelta(days=1)

        if last_task_end >= launch_date_obj:
            overrun_days = (last_task_end - launch_date_obj).days
            warning_msg = f"⚠️【時程衝突警告】工作將進行到 {last_task_end}，比上線日晚了 {overrun_days} 天。"

        if real_prep_end >= real_prep_start:
            schedule.append(
                {
                    "Task": "預備上線",
                    "Owner": "Ad2",
                    "Start Date": real_prep_start,
                    "End Date": real_prep_end,
                    "Type": "Prep",
                }
            )

        launch_end = add_workdays(launch_date_obj, launch_task_config["days"])
        schedule.append(
            {
                "Task": launch_task_config["task"],
                "Owner": launch_task_config["owner"],
                "Start Date": launch_date_obj,
                "End Date": launch_end,
                "Type": "Launch",
            }
        )

    df_schedule = pd.DataFrame(schedule)
    return df_schedule, warning_msg, holidays_dt


def build_excel_bytes(df_schedule, holidays_config, holidays_dt, launch_date_obj, collapse_threshold):
    output = io.BytesIO()

    min_date = df_schedule["Start Date"].min()
    max_date = df_schedule["End Date"].max()
    if launch_date_obj and launch_date_obj > max_date:
        max_date = launch_date_obj

    full_dates = list(pd.date_range(min_date, max_date, freq="D"))

    display_columns = []
    prep_task_row = df_schedule[df_schedule["Type"] == "Prep"]
    prep_task = None
    if not prep_task_row.empty:
        r = prep_task_row.iloc[0]
        prep_task = {
            "Start Date": r["Start Date"],
            "End Date": r["End Date"],
        }

    if prep_task and (prep_task["End Date"] - prep_task["Start Date"]).days + 1 >= collapse_threshold:
        keep_start = prep_task["Start Date"]
        resume_date = prep_task["End Date"] + timedelta(days=1)
        break_added = False

        for d in full_dates:
            d_date = d.date()
            if d_date >= resume_date or d_date <= keep_start:
                display_columns.append(d)
            else:
                if not break_added:
                    display_columns.append("BREAK")
                    break_added = True
    else:
        display_columns = full_dates

    def is_workday(d):
        return (d.weekday() < 5) and (d not in holidays_dt)

    holiday_blocks_info = []
    current_block_start = -1
    current_block_dates = []

    for i, col_item in enumerate(display_columns):
        is_holiday_day = False
        if col_item != "BREAK":
            d_date = col_item.date()
            if not is_workday(d_date):
                is_holiday_day = False if (launch_date_obj and d_date == launch_date_obj) else True

        if is_holiday_day:
            if current_block_start == -1:
                current_block_start = i
            current_block_dates.append(col_item.date())
        else:
            if current_block_start != -1:
                found_name = next(
                    (holidays_config[d.strftime("%Y-%m-%d")] for d in current_block_dates if d.strftime("%Y-%m-%d") in holidays_config),
                    None,
                )
                if found_name:
                    holiday_blocks_info.append(
                        {"start_col": current_block_start, "end_col": i - 1, "name": "\n".join(list(found_name))}
                    )
                elif len(current_block_dates) > 4:
                    holiday_blocks_info.append(
                        {"start_col": current_block_start, "end_col": i - 1, "name": "長\n假"}
                    )
            current_block_start, current_block_dates = -1, []

    if current_block_start != -1:
        found_name = next(
            (holidays_config[d.strftime("%Y-%m-%d")] for d in current_block_dates if d.strftime("%Y-%m-%d") in holidays_config),
            None,
        )
        if found_name:
            holiday_blocks_info.append(
                {"start_col": current_block_start, "end_col": len(display_columns) - 1, "name": "\n".join(list(found_name))}
            )

    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book
    worksheet = workbook.add_worksheet("時程表")

    font = "Microsoft JhengHei"
    FONT_SIZE = 11
    border_fmt = {"border": 1, "border_color": "#000000"}

    def F(**kwargs):
        base = {"font_name": font, "font_size": FONT_SIZE, **kwargs}
        return workbook.add_format(base)

    fmt_center = F(align="center", valign="vcenter", **border_fmt)
    fmt_left = F(align="left", valign="vcenter", **border_fmt)
    fmt_weekend = F(bg_color=COLOR_WEEKEND, align="center", valign="vcenter", **border_fmt)
    fmt_date_num = F(align="center", valign="vcenter", **border_fmt)
    fmt_holiday_merged = F(
        align="center",
        valign="vcenter",
        text_wrap=True,
        bg_color=COLOR_HOLIDAY_BG,
        border=1,
        font_color=COLOR_HOLIDAY_TEXT,
        bold=True,
    )
    fmt_header_main = F(bold=True, align="center", valign="vcenter", bg_color="#FFFFFF", **border_fmt)
    fmt_bar_client = F(bg_color=COLOR_CLIENT_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_bar_ad2 = F(bg_color=COLOR_AD2_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_bar_launch = F(bg_color=COLOR_LAUNCH_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_bar_prep = F(bg_color=COLOR_PREP_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_legend_client = F(bg_color=COLOR_LEGEND_CLIENT, align="center", valign="vcenter", **border_fmt)
    fmt_legend_ad2 = F(bg_color=COLOR_LEGEND_AD2, align="center", valign="vcenter", **border_fmt)
    fmt_break_merge = F(align="center", valign="vcenter", **border_fmt)

    worksheet.write(0, 2, "客戶", fmt_legend_client)
    worksheet.write(0, 3, "Ad2", fmt_legend_ad2)
    worksheet.merge_range(1, 0, 3, 0, "製作時程", fmt_header_main)
    worksheet.merge_range(1, 1, 3, 1, "Action by", fmt_header_main)

    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 12)

    col_start, row_start = 2, 4
    current_month, merge_start_col, month_color_idx = None, col_start, 0
    break_cols_excel = []

    for i, item in enumerate(display_columns):
        col = col_start + i
        if item == "BREAK":
            worksheet.set_column(col, col, 4)
            break_cols_excel.append(col)
            continue

        d = item
        if current_month is None:
            current_month = d.month

        if d.month != current_month:
            month_fmt = F(
                bold=True,
                align="center",
                valign="vcenter",
                bg_color=MONTH_COLORS[month_color_idx % len(MONTH_COLORS)],
                **border_fmt,
            )
            worksheet.merge_range(
                1,
                merge_start_col,
                1,
                col - 1,
                date(d.year, current_month, 1).strftime("%b").upper(),
                month_fmt,
            )
            current_month, merge_start_col, month_color_idx = d.month, col, month_color_idx + 1

        if i == len(display_columns) - 1:
            month_fmt = F(
                bold=True,
                align="center",
                valign="vcenter",
                bg_color=MONTH_COLORS[month_color_idx % len(MONTH_COLORS)],
                **border_fmt,
            )
            worksheet.merge_range(1, merge_start_col, 1, col, d.strftime("%b").upper(), month_fmt)

        is_h = not is_workday(d.date())
        weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}
        worksheet.write(2, col, d.day, fmt_weekend if is_h else fmt_date_num)
        worksheet.write(3, col, weekday_map[d.weekday()], fmt_weekend if is_h else fmt_date_num)
        worksheet.set_column(col, col, 4.5)

    last_task_row = row_start + len(df_schedule) - 1

    for c in break_cols_excel:
        worksheet.merge_range(2, c, last_task_row, c, "～", fmt_break_merge)

    for idx, item in df_schedule.iterrows():
        row = row_start + idx
        worksheet.write(row, 0, item["Task"], fmt_left)
        worksheet.write(row, 1, item["Owner"], fmt_left)

        if item["Type"] == "Launch":
            bar_fmt = fmt_bar_launch
        elif item["Type"] == "Prep":
            bar_fmt = fmt_bar_prep
        elif "客戶" in item["Owner"]:
            bar_fmt = fmt_bar_client
        else:
            bar_fmt = fmt_bar_ad2

        for i, col_item in enumerate(display_columns):
            if col_item == "BREAK":
                continue

            col = col_start + i
            d_date = col_item.date()
            if item["Start Date"] <= d_date <= item["End Date"]:
                if item["Type"] in ["Launch", "Prep"] or is_workday(d_date):
                    worksheet.write(row, col, "", bar_fmt)
                else:
                    worksheet.write(row, col, "", fmt_weekend)
            else:
                worksheet.write(row, col, "", fmt_weekend if not is_workday(d_date) else fmt_center)

    for block in holiday_blocks_info:
        c1, c2 = col_start + block["start_col"], col_start + block["end_col"]
        if c1 == c2:
            worksheet.merge_range(row_start, c1, last_task_row, c1, block["name"], fmt_holiday_merged)
        else:
            worksheet.merge_range(row_start, c1, last_task_row, c2, block["name"], fmt_holiday_merged)

    writer.close()
    output.seek(0)
    return output.getvalue()


def run_schedule(project_name, mode_display, start_date_value, launch_date_value, collapse_threshold, tasks_df, holidays_text):
    holidays = parse_holidays(holidays_text)
    tasks = normalize_tasks(tasks_df)
    calculation_mode = MODE_MAP[mode_display]

    start_date_obj = start_date_value if start_date_value else None
    launch_date_obj = launch_date_value if launch_date_value else None

    df_schedule, warning_msg, holidays_dt = build_scheduler(
        tasks_config=tasks,
        holidays_config=holidays,
        calculation_mode=calculation_mode,
        start_date_obj=start_date_obj,
        launch_date_obj=launch_date_obj,
    )

    excel_bytes = build_excel_bytes(
        df_schedule=df_schedule,
        holidays_config=holidays,
        holidays_dt=holidays_dt,
        launch_date_obj=launch_date_obj,
        collapse_threshold=int(collapse_threshold),
    )

    st.session_state.schedule_df = df_schedule
    st.session_state.warning_msg = warning_msg
    st.session_state.excel_bytes = excel_bytes
    st.session_state.project_name = project_name


def reset_defaults():
    st.session_state.tasks_df = pd.DataFrame(DEFAULT_TASKS)
    st.session_state.holidays_text = "\n".join([f"{k},{v}" for k, v in DEFAULT_HOLIDAYS.items()])
    st.session_state.schedule_df = None
    st.session_state.excel_bytes = None
    st.session_state.warning_msg = ""


# =========================
# 介面
# =========================
st.title("製作時程排程工具")
st.caption("請先填寫專案資訊，再調整流程內容；若某一步需固定在上線當天，請勾選「上線日」。")

with st.container(border=True):
    st.subheader("專案設定")

    col1, col2, col3 = st.columns([2, 1.3, 1.1])
    with col1:
        project_name = st.text_input("專案名稱", value=DEFAULT_PROJECT_NAME)
        mode_display = st.selectbox("排程方式", list(MODE_MAP.keys()), index=list(MODE_MAP.keys()).index(DEFAULT_MODE))
    with col2:
        start_date_value = st.date_input("開始日期", value=DEFAULT_START_DATE)
        launch_date_value = st.date_input("上線日期", value=DEFAULT_LAUNCH_DATE)
    with col3:
        collapse_threshold = st.number_input("日期縮略門檻", min_value=1, max_value=30, value=DEFAULT_COLLAPSE_THRESHOLD, step=1)
        st.markdown("")
        if st.button("重設為預設內容", use_container_width=True):
            reset_defaults()
            st.rerun()

# 按鈕往上移：放在任務表格前
action_col1, action_col2 = st.columns([1.2, 5])
with action_col1:
    generate = st.button("產出時程表", type="primary", use_container_width=True)

st.markdown("")

left, right = st.columns([3.2, 1.8])

with left:
    st.subheader("流程設定")
    st.session_state.tasks_df = st.data_editor(
        st.session_state.tasks_df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "任務名稱": st.column_config.TextColumn("任務名稱", required=True, width="large"),
            "負責人": st.column_config.SelectboxColumn("負責人", options=["Ad2", "客戶"], required=True, width="small"),
            "工作天數": st.column_config.NumberColumn("工作天數", min_value=1, max_value=365, step=1, required=True, width="small"),
            "上線日": st.column_config.CheckboxColumn("上線日", help="若此步驟需固定在上線當天，請勾選。", width="small"),
        },
        key="tasks_editor",
    )
    st.caption("可直接新增、刪除或修改任務。若未勾選任何一筆「上線日」，系統會自動將最後一筆視為上線日。")

with right:
    st.subheader("假日設定")
    st.session_state.holidays_text = st.text_area(
        "假日清單（每行一筆，格式：YYYY-MM-DD,名稱）",
        value=st.session_state.holidays_text,
        height=340,
    )
    st.caption("建議保留預設國定假日，再視需要補上公司內部休假日。")

if generate:
    try:
        run_schedule(
            project_name=project_name,
            mode_display=mode_display,
            start_date_value=start_date_value,
            launch_date_value=launch_date_value,
            collapse_threshold=collapse_threshold,
            tasks_df=st.session_state.tasks_df,
            holidays_text=st.session_state.holidays_text,
        )
        st.success("已完成時程表計算。")
    except Exception as e:
        st.error(f"產出失敗：{e}")

if st.session_state.warning_msg:
    st.warning(st.session_state.warning_msg)

if st.session_state.schedule_df is not None:
    st.subheader("排程預覽")
    preview_df = st.session_state.schedule_df.copy()
    preview_df["Start Date"] = preview_df["Start Date"].astype(str)
    preview_df["End Date"] = preview_df["End Date"].astype(str)
    preview_df = preview_df.rename(
        columns={
            "Task": "任務名稱",
            "Owner": "負責人",
            "Start Date": "開始日期",
            "End Date": "結束日期",
            "Type": "類別",
        }
    )
    preview_df["類別"] = preview_df["類別"].replace(
        {"Normal": "一般流程", "Launch": "上線任務", "Prep": "預備上線"}
    )

    st.dataframe(preview_df, use_container_width=True, hide_index=True)

    filename = f"{datetime.now().strftime('%m%d')}_{project_name}.xlsx"
    st.download_button(
        "下載 Excel 時程表",
        data=st.session_state.excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="secondary",
    )

with st.expander("使用建議"):
    st.markdown(
        """
- 目前不保留「匯入設定 JSON」，因為這版先以日常快速操作為主，畫面會更乾淨。
- 若你之後常常套用固定流程模板，再加回「範本匯入／匯出」會比較合理。
- 這版也拿掉了多餘的啟用勾選，避免每次都要多做一步。
        """
    )
