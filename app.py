
import io
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="製作時程排程工具",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================
# 基本設定
# =========================
DEFAULT_PROJECT_NAME = ""
MODE_OPTIONS = ["製作日推進", "上線日回推", "同時指定開始與上線日期"]
DEFAULT_MODE = MODE_OPTIONS[0]
DEFAULT_START_DATE = date.today()
DEFAULT_LAUNCH_DATE = date.today()
DEFAULT_COLLAPSE_THRESHOLD = 2

DEFAULT_TASKS = [
    {"顯示": True, "任務名稱": "提供素材", "Action By": "客戶", "工作天數": 1, "上線日": False},
    {"顯示": True, "任務名稱": "視覺製作", "Action By": "Ad2", "工作天數": 3, "上線日": False},
    {"顯示": True, "任務名稱": "客戶確認", "Action By": "客戶", "工作天數": 1, "上線日": False},
    {"顯示": True, "任務名稱": "視覺調整", "Action By": "Ad2", "工作天數": 2, "上線日": False},
    {"顯示": True, "任務名稱": "客戶確認", "Action By": "客戶", "工作天數": 1, "上線日": False},
    {"顯示": True, "任務名稱": "廣告進稿", "Action By": "Ad2", "工作天數": 1, "上線日": False},
    {"顯示": True, "任務名稱": "廣告上線", "Action By": "Ad2", "工作天數": 1, "上線日": True},
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
    "製作日推進": "forward",
    "上線日回推": "backward",
    "同時指定開始與上線日期": "double",
}

EXCEL_COLOR_CLIENT_BAR = '#EA9B56'
EXCEL_COLOR_AD2_BAR = '#4BACC6'
EXCEL_COLOR_LAUNCH_BAR = '#FF0000'
EXCEL_COLOR_PREP_BAR = '#92D050'
EXCEL_COLOR_WEEKEND = '#D9D9D9'
EXCEL_COLOR_HOLIDAY_TEXT = '#595959'
MONTH_COLORS = ['#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#E7E6E6']

UI_COLOR_BG = "#FAFAF8"
UI_COLOR_BORDER = "#ECE8E1"
UI_COLOR_TEXT = "#2F2A24"
UI_COLOR_MUTED = "#7A736A"
UI_COLOR_PRIMARY = "#C97B7B"
UI_COLOR_PRIMARY_HOVER = "#B86A6A"
UI_COLOR_AD2 = "#4BACC6"
UI_COLOR_CLIENT = "#EA9B56"
UI_COLOR_LAUNCH = "#D47B7B"
UI_COLOR_PREP = "#92D050"
UI_COLOR_WEEKEND = "#F2F2F2"

st.markdown(f"""
<style>
html, body, [data-testid="stAppViewContainer"] {{
  background: linear-gradient(180deg, #FCFBF9 0%, #F8F6F2 100%);
  color: {UI_COLOR_TEXT};
}}
[data-testid="stHeader"] {{ background: rgba(255,255,255,0); }}
.block-container {{
  padding-top: 2rem !important;
  padding-bottom: 3rem !important;
  max-width: 1400px;
}}
[data-testid="stSidebar"] {{
  background: #FBFAF8;
  border-right: 1px solid {UI_COLOR_BORDER};
}}
div.stButton > button[kind="primary"],
div.stDownloadButton > button[kind="primary"] {{
    background-color: {UI_COLOR_PRIMARY} !important;
    border: 1px solid {UI_COLOR_PRIMARY} !important;
    color: white !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    box-shadow: 0 6px 18px rgba(201,123,123,0.18);
}}
div.stButton > button[kind="primary"]:hover,
div.stDownloadButton > button[kind="primary"]:hover {{
    background-color: {UI_COLOR_PRIMARY_HOVER} !important;
    border-color: {UI_COLOR_PRIMARY_HOVER} !important;
}}
div.stButton > button[kind="secondary"] {{
    border-radius: 10px !important;
    border: 1px solid {UI_COLOR_BORDER} !important;
    background: white !important;
}}
[data-testid="stTextInputRootElement"],
[data-testid="stDateInputField"],
[data-baseweb="select"] > div,
[data-testid="stNumberInput"] input {{
    border-radius: 12px !important;
}}
[data-testid="stDataFrame"], [data-testid="stTable"] {{
    border-radius: 12px;
}}
.section-title {{
    font-size: 1.08rem;
    font-weight: 700;
    margin-bottom: 0.2rem;
}}
.section-sub {{
    color: {UI_COLOR_MUTED};
    font-size: 0.92rem;
    margin-bottom: 0.9rem;
}}
.preview-note {{
    color: {UI_COLOR_MUTED};
    font-size: 0.9rem;
    margin-top: -0.15rem;
    margin-bottom: 0.75rem;
}}
.info-chip-wrap {{
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    margin-bottom: 12px;
}}
.info-chip {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 10px;
    background: #F8F6F3;
    border: 1px solid {UI_COLOR_BORDER};
    border-radius: 999px;
    font-size: 12px;
    color: {UI_COLOR_MUTED};
}}
.gantt-wrap {{
    overflow-x: auto;
    border: 1px solid {UI_COLOR_BORDER};
    border-radius: 16px;
    background: #fff;
}}
.gantt-table {{
    border-collapse: collapse;
    width: max-content;
    min-width: 100%;
    font-size: 13px;
}}
.gantt-table th, .gantt-table td {{
    border: 1px solid #F0ECE6;
    text-align: center;
    padding: 0;
    height: 36px;
}}
.gantt-table .sticky-left {{
    position: sticky;
    left: 0;
    z-index: 3;
    background: #fff;
}}
.gantt-table .sticky-left-2 {{
    position: sticky;
    left: 190px;
    z-index: 3;
    background: #fff;
}}
.gantt-table .task-col {{
    width: 190px;
    min-width: 190px;
    max-width: 190px;
    padding: 0 12px;
    text-align: left;
    font-weight: 600;
}}
.gantt-table .owner-col {{
    width: 96px;
    min-width: 96px;
    max-width: 96px;
    padding: 0 8px;
    text-align: center;
}}
.gantt-table .month-row th {{
    background: #FBFAF8;
    font-weight: 700;
    height: 32px;
}}
.gantt-table .date-head {{
    width: 38px;
    min-width: 38px;
    max-width: 38px;
    line-height: 1.15;
    font-size: 11px;
    background: #fff;
}}
.gantt-table .weekend-head {{ background: #F4F4F4; }}
.gantt-table .break-head, .gantt-table .break-cell {{
    width: 26px;
    min-width: 26px;
    max-width: 26px;
    background: #F7F7F7;
    color: #777;
    font-weight: 700;
}}
.gantt-table .empty-cell {{ background: #fff; }}
.gantt-table .weekend-cell {{ background: #F3F3F3; }}
.gantt-table .bar-ad2 {{ background: {UI_COLOR_AD2}; }}
.gantt-table .bar-client {{ background: {UI_COLOR_CLIENT}; }}
.gantt-table .bar-launch {{ background: {UI_COLOR_LAUNCH}; }}
.gantt-table .bar-prep {{ background: {UI_COLOR_PREP}; }}
.legend {{
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin-bottom: 12px;
    font-size: 12px;
    color: {UI_COLOR_MUTED};
}}
.legend-item {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
}}
.legend-dot {{
    width: 12px;
    height: 12px;
    border-radius: 4px;
    display: inline-block;
}}
.small-gap {{ height: 0.35rem; }}
.large-gap {{ height: 1.8rem; }}

/* 隱藏 data editor 左側內建選取欄 */
[data-testid="stDataEditor"] [role="grid"] [aria-colindex="1"] {{
    display: none !important;
}}
[data-testid="stDataEditor"] [role="columnheader"][aria-colindex="1"] {{
    display: none !important;
}}
</style>
""", unsafe_allow_html=True)

def init_state():
    defaults = {
        "project_name": DEFAULT_PROJECT_NAME,
        "mode_display": DEFAULT_MODE,
        "start_date_value": DEFAULT_START_DATE,
        "launch_date_value": DEFAULT_LAUNCH_DATE,
        "collapse_threshold": DEFAULT_COLLAPSE_THRESHOLD,
        "tasks_df": pd.DataFrame(DEFAULT_TASKS),
        "holidays_text": "\n".join([f"{k},{v}" for k, v in DEFAULT_HOLIDAYS.items()]),
        "schedule_df": None,
        "excel_bytes": None,
        "warning_msg": "",
        "last_generated_name": DEFAULT_PROJECT_NAME,
        "display_columns": None,
        "holidays_dt": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

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
        pd.to_datetime(d)
        holidays[d] = name
    return holidays

def normalize_tasks(df: pd.DataFrame) -> list[dict]:
    if df is None or df.empty:
        raise ValueError("請至少保留一筆任務。")

    df = df.copy()
    required_cols = ["顯示", "任務名稱", "Action By", "工作天數", "上線日"]
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"缺少欄位：{c}")

    df["顯示"] = df["顯示"].fillna(True).astype(bool)
    df["任務名稱"] = df["任務名稱"].fillna("").astype(str).str.strip()
    df["Action By"] = df["Action By"].fillna("Ad2").astype(str).str.strip()
    df["工作天數"] = pd.to_numeric(df["工作天數"], errors="coerce")
    df["上線日"] = df["上線日"].fillna(False).astype(bool)

    df = df[(df["顯示"] == True) & (df["任務名稱"] != "")].copy()

    if df.empty:
        raise ValueError("目前沒有可排程的任務，請至少保留一筆顯示中的任務。")
    if (df["工作天數"].isna()).any() or (df["工作天數"] <= 0).any():
        raise ValueError("工作天數必須為大於 0 的整數。")

    launch_count = int(df["上線日"].sum())
    if launch_count > 1:
        raise ValueError("「上線日」只能勾選一筆。")
    if launch_count == 0:
        df.loc[df.index[-1], "上線日"] = True

    tasks = []
    for _, row in df.iterrows():
        tasks.append(
            {
                "task": row["任務名稱"],
                "owner": row["Action By"],
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
    launch_tasks = [t for t in tasks_config if t["is_launch"]]
    launch_task_config = launch_tasks[0] if launch_tasks else tasks_config[-1]

    if calculation_mode == "backward":
        if not launch_date_obj:
            raise ValueError("「上線日回推」需要填寫上線日期。")
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

            temp_schedule.append({
                "Task": t["task"],
                "Owner": t["owner"],
                "Start Date": start_date,
                "End Date": end_date,
                "Type": "Launch" if t["is_launch"] else "Normal",
            })
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
            schedule.append({
                "Task": t["task"],
                "Owner": t["owner"],
                "Start Date": start_d,
                "End Date": end_d,
                "Type": "Launch" if t["is_launch"] else "Normal",
            })
            prev_end = end_d
    else:
        if not start_date_obj or not launch_date_obj:
            raise ValueError("「同時指定開始與上線日期」需要同時填寫開始日期與上線日期。")
        curr_start = ensure_workday_forward(start_date_obj)
        prev_end = None
        normal_tasks = [t for t in tasks_config if not t["is_launch"]]

        for idx, t in enumerate(normal_tasks):
            start_d = curr_start if idx == 0 else get_next_workday(prev_end)
            end_d = add_workdays(start_d, t["days"])
            schedule.append({
                "Task": t["task"],
                "Owner": t["owner"],
                "Start Date": start_d,
                "End Date": end_d,
                "Type": "Normal",
            })
            prev_end = end_d

        last_task_end = schedule[-1]["End Date"] if schedule else start_date_obj
        real_prep_start = last_task_end + timedelta(days=1)
        real_prep_end = launch_date_obj - timedelta(days=1)

        if last_task_end >= launch_date_obj:
            overrun_days = (last_task_end - launch_date_obj).days
            warning_msg = f"⚠️【時程衝突警告】工作將進行到 {last_task_end}，比上線日晚了 {overrun_days} 天。"

        if real_prep_end >= real_prep_start:
            schedule.append({
                "Task": "預備上線",
                "Owner": "Ad2",
                "Start Date": real_prep_start,
                "End Date": real_prep_end,
                "Type": "Prep",
            })

        launch_end = add_workdays(launch_date_obj, launch_task_config["days"])
        schedule.append({
            "Task": launch_task_config["task"],
            "Owner": launch_task_config["owner"],
            "Start Date": launch_date_obj,
            "End Date": launch_end,
            "Type": "Launch",
        })

    return pd.DataFrame(schedule), warning_msg, holidays_dt

def prepare_display_columns(df_schedule, holidays_dt, launch_date_obj, collapse_threshold):
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
        prep_task = {"Start Date": r["Start Date"], "End Date": r["End Date"]}

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
    return display_columns

def build_excel_bytes(df_schedule, holidays_config, holidays_dt, launch_date_obj, collapse_threshold):
    output = io.BytesIO()
    display_columns = prepare_display_columns(df_schedule, holidays_dt, launch_date_obj, collapse_threshold)

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
                    holiday_blocks_info.append({"start_col": current_block_start, "end_col": i - 1, "name": "\n".join(list(found_name))})
                elif len(current_block_dates) > 4:
                    holiday_blocks_info.append({"start_col": current_block_start, "end_col": i - 1, "name": "長\n假"})
            current_block_start, current_block_dates = -1, []

    if current_block_start != -1:
        found_name = next(
            (holidays_config[d.strftime("%Y-%m-%d")] for d in current_block_dates if d.strftime("%Y-%m-%d") in holidays_config),
            None,
        )
        if found_name:
            holiday_blocks_info.append({"start_col": current_block_start, "end_col": len(display_columns) - 1, "name": "\n".join(list(found_name))})

    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book
    worksheet = workbook.add_worksheet("時程表")
    font = "Microsoft JhengHei"
    FONT_SIZE = 11
    border_fmt = {"border": 1, "border_color": "#000000"}

    def F(**kwargs):
        return workbook.add_format({"font_name": font, "font_size": FONT_SIZE, **kwargs})

    fmt_center = F(align="center", valign="vcenter", **border_fmt)
    fmt_left = F(align="left", valign="vcenter", **border_fmt)
    fmt_weekend = F(bg_color=EXCEL_COLOR_WEEKEND, align="center", valign="vcenter", **border_fmt)
    fmt_date_num = F(align="center", valign="vcenter", **border_fmt)
    fmt_holiday_merged = F(align="center", valign="vcenter", text_wrap=True, bg_color=EXCEL_COLOR_WEEKEND, border=1, font_color=EXCEL_COLOR_HOLIDAY_TEXT, bold=True)
    fmt_header_main = F(bold=True, align="center", valign="vcenter", bg_color="#FFFFFF", **border_fmt)
    fmt_bar_client = F(bg_color=EXCEL_COLOR_CLIENT_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_bar_ad2 = F(bg_color=EXCEL_COLOR_AD2_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_bar_launch = F(bg_color=EXCEL_COLOR_LAUNCH_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_bar_prep = F(bg_color=EXCEL_COLOR_PREP_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_legend_client = F(bg_color=EXCEL_COLOR_CLIENT_BAR, align="center", valign="vcenter", **border_fmt)
    fmt_legend_ad2 = F(bg_color=EXCEL_COLOR_AD2_BAR, align="center", valign="vcenter", **border_fmt)
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
            month_fmt = F(bold=True, align="center", valign="vcenter", bg_color=MONTH_COLORS[month_color_idx % len(MONTH_COLORS)], **border_fmt)
            worksheet.merge_range(1, merge_start_col, 1, col - 1, date(d.year, current_month, 1).strftime("%b").upper(), month_fmt)
            current_month, merge_start_col, month_color_idx = d.month, col, month_color_idx + 1

        if i == len(display_columns) - 1:
            month_fmt = F(bold=True, align="center", valign="vcenter", bg_color=MONTH_COLORS[month_color_idx % len(MONTH_COLORS)], **border_fmt)
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
    return output.getvalue(), display_columns

def get_effective_dates(mode_display):
    if mode_display == "製作日推進":
        return st.session_state.start_date_value, None
    if mode_display == "上線日回推":
        return None, st.session_state.launch_date_value
    return st.session_state.start_date_value, st.session_state.launch_date_value

def render_gantt_html(df_schedule, display_columns, holidays_dt):
    def is_workday(d):
        return (d.weekday() < 5) and (d not in holidays_dt)

    weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}

    month_cells = []
    i = 0
    while i < len(display_columns):
        item = display_columns[i]
        if item == "BREAK":
            month_cells.append('<th class="break-head" rowspan="2">～</th>')
            i += 1
            continue
        month = item.strftime("%m月")
        span = 1
        j = i + 1
        while j < len(display_columns):
            nxt = display_columns[j]
            if nxt == "BREAK" or nxt.strftime("%m") != item.strftime("%m"):
                break
            span += 1
            j += 1
        month_cells.append(f'<th colspan="{span}">{month}</th>')
        i = j

    header_row = []
    for item in display_columns:
        if item == "BREAK":
            header_row.append('<th class="break-head">～</th>')
        else:
            d = item.date()
            cls = "date-head weekend-head" if not is_workday(d) else "date-head"
            header_row.append(f'<th class="{cls}">{item.strftime("%m/%d")}<br>{weekday_map[item.weekday()]}</th>')

    body_rows = []
    for _, row in df_schedule.iterrows():
        cells = [
            f'<td class="task-col sticky-left">{row["Task"]}</td>',
            f'<td class="owner-col sticky-left-2">{row["Owner"]}</td>',
        ]
        for item in display_columns:
            if item == "BREAK":
                cells.append('<td class="break-cell">～</td>')
                continue

            d = item.date()
            base_cls = "weekend-cell" if not is_workday(d) else "empty-cell"

            if row["Start Date"] <= d <= row["End Date"]:
                if row["Type"] == "Launch":
                    cls = "bar-launch"
                elif row["Type"] == "Prep":
                    cls = "bar-prep"
                elif "客戶" in row["Owner"]:
                    cls = "bar-client"
                else:
                    cls = "bar-ad2"
                cells.append(f'<td class="{cls}"></td>')
            else:
                cells.append(f'<td class="{base_cls}"></td>')

        body_rows.append(f"<tr>{''.join(cells)}</tr>")

    html = f"""
    <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:{UI_COLOR_AD2};"></span>Ad2</span>
        <span class="legend-item"><span class="legend-dot" style="background:{UI_COLOR_CLIENT};"></span>客戶</span>
        <span class="legend-item"><span class="legend-dot" style="background:{UI_COLOR_LAUNCH};"></span>上線</span>
        <span class="legend-item"><span class="legend-dot" style="background:{UI_COLOR_PREP};"></span>預備上線</span>
        <span class="legend-item"><span class="legend-dot" style="background:{UI_COLOR_WEEKEND};"></span>假日／週末</span>
    </div>
    <div class="gantt-wrap">
        <table class="gantt-table">
            <thead>
                <tr class="month-row">
                    <th class="task-col sticky-left" rowspan="2">任務名稱</th>
                    <th class="owner-col sticky-left-2" rowspan="2">Action By</th>
                    {''.join(month_cells)}
                </tr>
                <tr>
                    {''.join(header_row)}
                </tr>
            </thead>
            <tbody>
                {''.join(body_rows)}
            </tbody>
        </table>
    </div>
    """
    return html

def generate_schedule():
    holidays = parse_holidays(st.session_state.holidays_text)
    tasks = normalize_tasks(st.session_state.tasks_df)
    calculation_mode = MODE_MAP[st.session_state.mode_display]
    start_date_obj, launch_date_obj = get_effective_dates(st.session_state.mode_display)

    df_schedule, warning_msg, holidays_dt = build_scheduler(
        tasks_config=tasks,
        holidays_config=holidays,
        calculation_mode=calculation_mode,
        start_date_obj=start_date_obj,
        launch_date_obj=launch_date_obj,
    )
    excel_bytes, display_columns = build_excel_bytes(
        df_schedule=df_schedule,
        holidays_config=holidays,
        holidays_dt=holidays_dt,
        launch_date_obj=launch_date_obj,
        collapse_threshold=int(st.session_state.collapse_threshold),
    )

    st.session_state.schedule_df = df_schedule
    st.session_state.warning_msg = warning_msg
    st.session_state.excel_bytes = excel_bytes
    st.session_state.last_generated_name = st.session_state.project_name or "未命名專案"
    st.session_state.display_columns = display_columns
    st.session_state.holidays_dt = holidays_dt

def reset_defaults():
    st.session_state.project_name = DEFAULT_PROJECT_NAME
    st.session_state.mode_display = DEFAULT_MODE
    st.session_state.start_date_value = DEFAULT_START_DATE
    st.session_state.launch_date_value = DEFAULT_LAUNCH_DATE
    st.session_state.collapse_threshold = DEFAULT_COLLAPSE_THRESHOLD
    st.session_state.tasks_df = pd.DataFrame(DEFAULT_TASKS)
    st.session_state.schedule_df = None
    st.session_state.excel_bytes = None
    st.session_state.warning_msg = ""
    st.session_state.last_generated_name = DEFAULT_PROJECT_NAME
    st.session_state.display_columns = None
    st.session_state.holidays_dt = None

st.title("製作時程排程工具")
st.caption("快速設定專案日期與流程後，即可產出 Excel 時程表；下方預覽會用色塊顯示整體節奏。")

with st.sidebar:
    st.subheader("假日設定")
    st.caption("可從左上角展開側邊欄，平常可收起不顯示。")
    st.text_area("假日清單（每行一筆，格式：YYYY-MM-DD,名稱）", key="holidays_text", height=420)
    st.caption("建議保留預設國定假日，再視需要補上公司內部休假日。")

with st.container(border=True):
    header_col1, header_col2 = st.columns([5.2, 1.1])
    with header_col1:
        st.markdown('<div class="section-title">專案設定</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">先決定排程方式與日期，再按下「產出時程表」。</div>', unsafe_allow_html=True)
    with header_col2:
        st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)
        st.button("重設", use_container_width=True, on_click=reset_defaults)

    row1_col1, row1_col2, row1_col3 = st.columns([2.5, 1.6, 1.0])
    with row1_col1:
        st.text_input("專案名稱", key="project_name", placeholder="請輸入專案名稱")
    with row1_col2:
        st.selectbox("排程方式", MODE_OPTIONS, key="mode_display")
    with row1_col3:
        st.number_input("日期縮略門檻", min_value=1, max_value=30, step=1, key="collapse_threshold")

    current_mode = st.session_state.mode_display
    start_disabled = current_mode == "上線日回推"
    launch_disabled = current_mode == "製作日推進"

    row2_col1, row2_col2, row2_col3 = st.columns([1.5, 1.5, 1.1])
    with row2_col1:
        st.date_input("開始日期", key="start_date_value", disabled=start_disabled, help="在「上線日回推」模式下，此欄位不需填寫。")
    with row2_col2:
        st.date_input("上線日期", key="launch_date_value", disabled=launch_disabled, help="在「製作日推進」模式下，此欄位不需填寫。")
    with row2_col3:
        st.markdown('<div class="large-gap"></div>', unsafe_allow_html=True)
        st.button("產出時程表", type="primary", use_container_width=True, on_click=generate_schedule)

st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)

if st.session_state.warning_msg:
    st.warning(st.session_state.warning_msg)

if st.session_state.schedule_df is not None:
    with st.container(border=True):
        preview_header_col, download_col = st.columns([5.2, 1.25])
        with preview_header_col:
            st.markdown('<div class="section-title">排程預覽</div>', unsafe_allow_html=True)
            st.markdown("""
            <div class="info-chip-wrap">
                <span class="info-chip">可左右滑動查看完整日期</span>
            </div>
            """, unsafe_allow_html=True)
        with download_col:
            filename = f"{datetime.now().strftime('%m%d')}_{st.session_state.last_generated_name}.xlsx"
            st.markdown('<div class="large-gap"></div>', unsafe_allow_html=True)
            st.download_button(
                "下載 Excel",
                data=st.session_state.excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        gantt_html = render_gantt_html(
            st.session_state.schedule_df,
            st.session_state.display_columns,
            st.session_state.holidays_dt,
        )
        st.markdown(gantt_html, unsafe_allow_html=True)

st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)

with st.container(border=True):
    st.markdown('<div class="section-title">流程設定</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">取消勾選「顯示」即可暫時隱藏該項目，不會進入排程。</div>', unsafe_allow_html=True)

    st.session_state.tasks_df = st.data_editor(
        st.session_state.tasks_df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_order=["顯示", "任務名稱", "Action By", "工作天數", "上線日"],
        column_config={
            "顯示": st.column_config.CheckboxColumn("顯示", help="取消勾選即可暫時隱藏該任務。", width="small"),
            "任務名稱": st.column_config.TextColumn("任務名稱", required=True, width="large"),
            "Action By": st.column_config.SelectboxColumn("Action By", options=["Ad2", "客戶"], required=True, width="medium"),
            "工作天數": st.column_config.NumberColumn("工作天數", min_value=1, max_value=365, step=1, required=True, width="small"),
            "上線日": st.column_config.CheckboxColumn("上線日", help="若此步驟需固定在上線當天，請勾選。", width="small"),
        },
        key="tasks_editor_v10",
    )
