
import io
import uuid
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="製作時程排程工具", page_icon="📅", layout="wide", initial_sidebar_state="collapsed")

# =========================
# 基本設定
# =========================
MODE_OPTIONS = ["製作日推進", "上線日回推", "同時指定開始與上線日期"]
MODE_MAP = {
    "製作日推進": "forward",
    "上線日回推": "backward",
    "同時指定開始與上線日期": "double",
}

DEFAULT_TASKS = [
    {"id": "task_1", "顯示": True, "任務名稱": "提供素材", "Action By": "客戶", "工作天數": 1, "上線日": False},
    {"id": "task_2", "顯示": True, "任務名稱": "視覺製作", "Action By": "Ad2", "工作天數": 3, "上線日": False},
    {"id": "task_3", "顯示": True, "任務名稱": "客戶確認", "Action By": "客戶", "工作天數": 1, "上線日": False},
    {"id": "task_4", "顯示": True, "任務名稱": "視覺調整", "Action By": "Ad2", "工作天數": 2, "上線日": False},
    {"id": "task_5", "顯示": True, "任務名稱": "客戶確認", "Action By": "客戶", "工作天數": 1, "上線日": False},
    {"id": "task_6", "顯示": True, "任務名稱": "廣告進稿", "Action By": "Ad2", "工作天數": 1, "上線日": False},
    {"id": "task_7", "顯示": True, "任務名稱": "廣告上線", "Action By": "Ad2", "工作天數": 1, "上線日": True},
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

# Excel colors
EXCEL_COLOR_CLIENT_BAR = '#EA9B56'
EXCEL_COLOR_AD2_BAR = '#4BACC6'
EXCEL_COLOR_LAUNCH_BAR = '#FF0000'
EXCEL_COLOR_PREP_BAR = '#92D050'
EXCEL_COLOR_WEEKEND = '#D9D9D9'
EXCEL_COLOR_HOLIDAY_TEXT = '#595959'
MONTH_COLORS = ['#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#E7E6E6']

# UI colors
UI_PRIMARY = "#C97B7B"
UI_PRIMARY_HOVER = "#B86A6A"
UI_BORDER = "#ECE8E1"
UI_MUTED = "#7A736A"
UI_AD2 = "#4BACC6"
UI_CLIENT = "#EA9B56"
UI_LAUNCH = "#D47B7B"
UI_PREP = "#92D050"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700;800&display=swap');

html, body, [data-testid="stAppViewContainer"] {{
  background: #F5F3EF;
  font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif;
}}
.block-container {{
  max-width: 1480px;
  padding-top: 2rem !important;
  padding-bottom: 4rem !important;
}}
[data-testid="stSidebar"] {{
  background: #FDFCFA;
  border-right: 1px solid #E5E1DA;
}}

/* ── Page title ── */
h1 {{
  font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif !important;
  font-size: 1.75rem !important;
  font-weight: 800 !important;
  color: #1E1B18 !important;
  letter-spacing: -0.3px !important;
  margin-bottom: 0 !important;
}}

/* ── Primary buttons ── */
div.stButton > button[kind="primary"],
div.stDownloadButton > button[kind="primary"] {{
  background: {UI_PRIMARY} !important;
  border: none !important;
  color: #fff !important;
  border-radius: 8px !important;
  font-weight: 700 !important;
  font-size: 0.88rem !important;
  letter-spacing: 0.3px !important;
  padding: 0.45rem 1rem !important;
  box-shadow: 0 2px 10px rgba(201,123,123,0.35) !important;
  transition: transform 0.12s, box-shadow 0.12s !important;
}}
div.stButton > button[kind="primary"]:hover,
div.stDownloadButton > button[kind="primary"]:hover {{
  background: {UI_PRIMARY_HOVER} !important;
  box-shadow: 0 4px 16px rgba(201,123,123,0.45) !important;
  transform: translateY(-1px) !important;
}}

/* ── Secondary / plain buttons ── */
div.stButton > button:not([kind="primary"]) {{
  border-radius: 7px !important;
  font-size: 0.83rem !important;
  border: 1px solid #DDD9D2 !important;
  color: #5A5550 !important;
  background: #FDFCFA !important;
  transition: background 0.1s, border-color 0.1s !important;
}}
div.stButton > button:not([kind="primary"]):hover {{
  background: #F5F2ED !important;
  border-color: #C5C0B8 !important;
}}

/* ── Section headers ── */
.section-title {{
  font-size: 1rem;
  font-weight: 700;
  color: #1E1B18;
  margin-bottom: 0.15rem;
  letter-spacing: 0.1px;
}}
.section-sub {{
  color: #8C8680;
  font-size: 0.845rem;
  margin-bottom: 0.85rem;
  line-height: 1.5;
}}

/* ── Containers ── */
[data-testid="stVerticalBlock"] > [data-testid="element-container"] > div[style*="border"] {{
  border-radius: 14px !important;
  border-color: #E5E1DA !important;
}}

/* ── Timeline wrapper ── */
.timeline-wrap {{
  overflow-x: auto;
  border: 1px solid #E5E1DA;
  border-radius: 14px;
  background: #fff;
  box-shadow: 0 2px 16px rgba(0,0,0,0.07);
  margin-top: 6px;
}}

/* ── Timeline table base ── */
.timeline-table {{
  border-collapse: collapse;
  width: max-content;
  min-width: 100%;
  font-size: 12.5px;
  font-family: 'Noto Sans TC', 'Microsoft JhengHei', sans-serif;
}}
.timeline-table th,
.timeline-table td {{
  border: 1px solid #EDE9E2;
  text-align: center;
  padding: 0;
  height: 34px;
}}

/* ── Month header row ── */
.timeline-table .month-row th {{
  height: 26px;
  background: #F0EDE7;
  font-weight: 700;
  font-size: 11.5px;
  color: #4A4540;
  letter-spacing: 0.6px;
}}

/* ── Date & weekday header ── */
.timeline-table .date-head {{
  width: 34px; min-width: 34px; max-width: 34px;
  font-size: 11px;
  line-height: 1.2;
  color: #4E4A46;
  background: #F8F6F2;
}}
.timeline-table .weekend-head {{
  background: #EEECE8 !important;
  color: #9A9590 !important;
}}
.timeline-table .weekend-cell {{
  background: #F4F2EE;
}}
.timeline-table .empty-cell {{
  background: #fff;
}}

/* ── Sticky left columns ── */
.timeline-table .task-col {{
  min-width: 186px; max-width: 186px; width: 186px;
  text-align: left; padding: 0 13px;
  font-weight: 600; font-size: 12.5px;
  background: #fff;
  position: sticky; left: 0; z-index: 3;
  border-right: 2px solid #DDD9D2;
  color: #1E1B18;
}}
.timeline-table .owner-col {{
  min-width: 90px; max-width: 90px; width: 90px;
  background: #fff;
  position: sticky; left: 186px; z-index: 3;
  font-size: 12px;
  color: #5A5550;
  border-right: 2px solid #DDD9D2;
}}

/* Sticky header cells */
.timeline-table .month-row .task-col,
.timeline-table .month-row .owner-col,
tr:nth-child(2) .task-col,
tr:nth-child(2) .owner-col,
tr:nth-child(3) .task-col,
tr:nth-child(3) .owner-col {{
  background: #F0EDE7;
}}

/* ── BREAK column ── */
.timeline-table .break-cell {{
  width: 22px; min-width: 22px; max-width: 22px;
  background: linear-gradient(180deg, #E8E4DC 0%, #D8D4CC 100%);
  color: #888078;
  font-weight: 800;
  font-size: 12px;
  writing-mode: vertical-rl;
  text-orientation: mixed;
  letter-spacing: 3px;
  vertical-align: middle;
  border-left: 2px solid #C8C4BC;
  border-right: 2px solid #C8C4BC;
}}

/* ── Bar cells ── */
.timeline-table .bar-ad2    {{ background: {UI_AD2};    border-color: #3A9CB6; }}
.timeline-table .bar-client {{ background: {UI_CLIENT}; border-color: #D88A46; }}
.timeline-table .bar-launch {{ background: {UI_LAUNCH}; border-color: #C46A6A; }}
.timeline-table .bar-prep   {{ background: {UI_PREP};   border-color: #7DC040; }}

/* ── Legend ── */
.legend {{
  display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 10px;
  font-size: 12px; color: {UI_MUTED};
  align-items: center;
  padding: 6px 2px;
}}
.legend-item {{ display: inline-flex; align-items: center; gap: 5px; }}
.legend-dot {{
  width: 11px; height: 11px; border-radius: 3px; display: inline-block;
  box-shadow: 0 1px 3px rgba(0,0,0,0.15);
}}

/* ── Task config table ── */
.task-config-area {{
  margin-top: 0.3rem;
}}
.task-config-head {{
  background: #F7F4EF;
  border: 1px solid #E6E0D7;
  border-radius: 10px;
  padding: 0.36rem 0.7rem;
  margin-bottom: 0.2rem;
}}
.task-head-label {{
  font-size: 11px;
  font-weight: 700;
  color: #8C8680;
  letter-spacing: 0.2px;
  padding: 0;
  text-align: center;
}}
.task-table-row {{
  padding: 0.2rem 0.15rem 0.12rem 0.15rem;
  border-bottom: 1px solid #ECE6DD;
}}
.task-table-row:last-child {{
  border-bottom: none;
}}

/* ── Op buttons ── */
.op-btn button {{
  font-size: 13px !important;
  padding: 0 !important;
  height: 2rem !important;
  min-height: 2rem !important;
  border-radius: 7px !important;
}}

/* ── Gaps ── */
.small-gap {{ height: 0.3rem; }}
.large-gap {{ height: 1.6rem; }}

/* ── Streamlit form input tweaks ── */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {{
  border-radius: 7px !important;
  border-color: #DDD9D2 !important;
  font-size: 0.875rem !important;
}}
[data-testid="stSelectbox"] > div {{
  border-radius: 7px !important;
}}
[data-testid="stDateInput"] input {{
  border-radius: 7px !important;
  border-color: #DDD9D2 !important;
}}

/* ── Compact flow section ── */
.flow-config-scope [data-testid="element-container"] {{
  margin-bottom: 0.04rem !important;
}}
.flow-config-scope [data-testid="stVerticalBlock"] {{
  gap: 0.02rem !important;
}}
.flow-config-scope label[data-testid="stWidgetLabel"] {{
  display: none !important;
}}
.flow-config-scope [data-testid="stTextInput"] input,
.flow-config-scope [data-testid="stNumberInput"] input,
.flow-config-scope [data-testid="stSelectbox"] > div,
.flow-config-scope [data-testid="stSelectbox"] [data-baseweb="select"],
.flow-config-scope [data-testid="stSelectbox"] input {{
  min-height: 2rem !important;
}}
.flow-config-scope [data-testid="stTextInput"] input,
.flow-config-scope [data-testid="stNumberInput"] input {{
  padding-top: 0.35rem !important;
  padding-bottom: 0.35rem !important;
}}
.flow-config-scope [data-testid="stCheckbox"] {{
  display: flex;
  justify-content: center;
}}
.flow-config-scope [data-testid="column"] {{
  align-items: center;
}}
</style>
""", unsafe_allow_html=True)

# =========================
# state
# =========================
def init_state():
    if "project_name" not in st.session_state:
        st.session_state.project_name = ""
    if "mode_display" not in st.session_state:
        st.session_state.mode_display = MODE_OPTIONS[0]
    if "start_date_value" not in st.session_state:
        st.session_state.start_date_value = date.today()
    if "launch_date_value" not in st.session_state:
        st.session_state.launch_date_value = date.today()
    if "collapse_threshold" not in st.session_state:
        st.session_state.collapse_threshold = 2
    if "tasks" not in st.session_state:
        st.session_state.tasks = [row.copy() for row in DEFAULT_TASKS]
    if "holidays_text" not in st.session_state:
        st.session_state.holidays_text = "\n".join(f"{k},{v}" for k, v in DEFAULT_HOLIDAYS.items())
    if "schedule_df" not in st.session_state:
        st.session_state.schedule_df = None
    if "excel_bytes" not in st.session_state:
        st.session_state.excel_bytes = None
    if "display_columns" not in st.session_state:
        st.session_state.display_columns = None
    if "holidays_dt" not in st.session_state:
        st.session_state.holidays_dt = None
    if "warning_msg" not in st.session_state:
        st.session_state.warning_msg = ""
    if "last_generated_name" not in st.session_state:
        st.session_state.last_generated_name = "未命名專案"
    if "status_msg" not in st.session_state:
        st.session_state.status_msg = ""

init_state()

def ensure_task_ids():
    tasks = st.session_state.tasks
    existing = set()
    for i, row in enumerate(tasks):
        rid = row.get("id")
        if not rid or rid in existing:
            row["id"] = f"task_{i+1}_{uuid.uuid4().hex[:6]}"
        existing.add(row["id"])

ensure_task_ids()

# =========================
# helpers
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
        pd.to_datetime(d.strip())
        holidays[d.strip()] = name.strip()
    return holidays

def get_active_tasks():
    tasks = []
    visible_rows = [row for row in st.session_state.tasks if row.get("顯示", True) and str(row.get("任務名稱", "")).strip()]
    if not visible_rows:
        raise ValueError("目前沒有可排程的任務，請至少保留一筆顯示中的任務。")

    launch_count = sum(bool(r.get("上線日", False)) for r in visible_rows)
    if launch_count > 1:
        raise ValueError("「上線日」只能勾選一筆。")
    if launch_count == 0:
        visible_rows[-1]["上線日"] = True

    for row in visible_rows:
        days = int(row.get("工作天數", 0) or 0)
        if days <= 0:
            raise ValueError(f"任務「{row.get('任務名稱','未命名')}」的工作天數需大於 0。")
        tasks.append({
            "task": str(row.get("任務名稱", "")).strip(),
            "owner": str(row.get("Action By", "Ad2")).strip() or "Ad2",
            "days": days,
            "is_launch": bool(row.get("上線日", False)),
        })
    return tasks

def build_scheduler(tasks_config, holidays_config, calculation_mode, start_date_obj, launch_date_obj):
    holidays_dt = [pd.to_datetime(h).date() for h in holidays_config.keys()]

    def is_workday(d):
        return (d.weekday() < 5) and (d not in holidays_dt)

    def subtract_workdays(start_date, days):
        current = start_date
        check_days = max(days - 1, 0)
        while check_days > 0:
            current -= timedelta(days=1)
            if is_workday(current):
                check_days -= 1
        return current

    def add_workdays(start_date, days):
        current = start_date
        check_days = max(days - 1, 0)
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
    launch_task_config = next((t for t in tasks_config if t["is_launch"]), tasks_config[-1])

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
                "Task": t["task"], "Owner": t["owner"],
                "Start Date": start_date, "End Date": end_date,
                "Type": "Launch" if t["is_launch"] else "Normal"
            })
        schedule = temp_schedule[::-1]

    elif calculation_mode == "forward":
        curr_start = ensure_workday_forward(start_date_obj or date.today())
        prev_end = None
        for idx, t in enumerate(tasks_config):
            start_d = launch_date_obj if (t["is_launch"] and launch_date_obj) else (curr_start if idx == 0 else get_next_workday(prev_end))
            end_d = add_workdays(start_d, t["days"])
            schedule.append({
                "Task": t["task"], "Owner": t["owner"],
                "Start Date": start_d, "End Date": end_d,
                "Type": "Launch" if t["is_launch"] else "Normal"
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
                "Task": t["task"], "Owner": t["owner"],
                "Start Date": start_d, "End Date": end_d, "Type": "Normal"
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
                "Task": "預備上線", "Owner": "Ad2",
                "Start Date": real_prep_start, "End Date": real_prep_end, "Type": "Prep"
            })

        launch_end = add_workdays(launch_date_obj, launch_task_config["days"])
        schedule.append({
            "Task": launch_task_config["task"], "Owner": launch_task_config["owner"],
            "Start Date": launch_date_obj, "End Date": launch_end, "Type": "Launch"
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

def compute_month_segments(display_columns, col_start):
    segments = []
    segment_start_col = None
    segment_month = None
    segment_year = None

    for i, item in enumerate(display_columns):
        if item == "BREAK":
            if segment_start_col is not None:
                segments.append((segment_start_col, col_start + i - 1, segment_year, segment_month))
                segment_start_col = None
                segment_month = None
                segment_year = None
            continue

        d = item
        excel_col = col_start + i
        if segment_start_col is None:
            segment_start_col = excel_col
            segment_month = d.month
            segment_year = d.year
        elif d.month != segment_month or d.year != segment_year:
            segments.append((segment_start_col, excel_col - 1, segment_year, segment_month))
            segment_start_col = excel_col
            segment_month = d.month
            segment_year = d.year

    if segment_start_col is not None:
        last_real_col = None
        for i in range(len(display_columns) - 1, -1, -1):
            if display_columns[i] != "BREAK":
                last_real_col = col_start + i
                break
        if last_real_col is not None:
            segments.append((segment_start_col, last_real_col, segment_year, segment_month))

    return segments

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
                found_name = next((holidays_config[d.strftime("%Y-%m-%d")] for d in current_block_dates if d.strftime("%Y-%m-%d") in holidays_config), None)
                if found_name:
                    holiday_blocks_info.append({"start_col": current_block_start, "end_col": i - 1, "name": "\n".join(list(found_name))})
                elif len(current_block_dates) > 4:
                    holiday_blocks_info.append({"start_col": current_block_start, "end_col": i - 1, "name": "長\n假"})
            current_block_start, current_block_dates = -1, []
    if current_block_start != -1:
        found_name = next((holidays_config[d.strftime("%Y-%m-%d")] for d in current_block_dates if d.strftime("%Y-%m-%d") in holidays_config), None)
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
    break_cols_excel = []

    month_segments = compute_month_segments(display_columns, col_start)
    month_color_idx = 0
    for start_col, end_col, year, month in month_segments:
        month_fmt = F(bold=True, align="center", valign="vcenter", bg_color=MONTH_COLORS[month_color_idx % len(MONTH_COLORS)], **border_fmt)
        month_label = date(year, month, 1).strftime("%b").upper()
        if start_col == end_col:
            worksheet.write(1, start_col, month_label, month_fmt)
        else:
            worksheet.merge_range(1, start_col, 1, end_col, month_label, month_fmt)
        month_color_idx += 1

    for i, item in enumerate(display_columns):
        col = col_start + i
        if item == "BREAK":
            worksheet.set_column(col, col, 4)
            break_cols_excel.append(col)
            continue
        d = item
        is_h = not is_workday(d.date())
        weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}
        worksheet.write(2, col, d.day, fmt_weekend if is_h else fmt_date_num)
        worksheet.write(3, col, weekday_map[d.weekday()], fmt_weekend if is_h else fmt_date_num)
        worksheet.set_column(col, col, 4.5)

    last_task_row = row_start + len(df_schedule) - 1
    for c in break_cols_excel:
        worksheet.merge_range(1, c, last_task_row, c, "～", fmt_break_merge)

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

def render_stable_preview(df_schedule, display_columns, holidays_dt):
    def is_workday(d):
        return (d.weekday() < 5) and (d not in holidays_dt)
    weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}

    # 3 header rows (month / date / weekday) + all task rows
    total_rowspan = 3 + len(df_schedule)

    month_cells = []
    i = 0
    while i < len(display_columns):
        item = display_columns[i]
        if item == "BREAK":
            # Merge entire column (including month/date/weekday headers + all task rows)
            month_cells.append(f'<th class="break-cell" rowspan="{total_rowspan}">～</th>')
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

    date_cells = []
    weekday_cells = []
    for item in display_columns:
        if item == "BREAK":
            continue  # already covered by rowspan above
        d = item.date()
        cls = "date-head weekend-head" if not is_workday(d) else "date-head"
        date_cells.append(f'<th class="{cls}">{d.day}</th>')
        weekday_cells.append(f'<th class="{cls}">{weekday_map[item.weekday()]}</th>')

    rows = []
    rows.append(
        "<tr class='month-row'>"
        "<th class='task-col'>任務名稱</th>"
        "<th class='owner-col'>Action By</th>"
        + "".join(month_cells) + "</tr>"
    )
    rows.append("<tr><th class='task-col'></th><th class='owner-col'></th>" + "".join(date_cells) + "</tr>")
    rows.append("<tr><th class='task-col'></th><th class='owner-col'></th>" + "".join(weekday_cells) + "</tr>")

    for _, row in df_schedule.iterrows():
        cells = [
            f'<td class="task-col">{row["Task"]}</td>',
            f'<td class="owner-col">{row["Owner"]}</td>',
        ]
        for item in display_columns:
            if item == "BREAK":
                continue  # already merged by rowspan — skip this cell
            d = item.date()
            base_cls = "weekend-cell" if not is_workday(d) else "empty-cell"
            if row["Start Date"] <= d <= row["End Date"]:
                if row["Type"] in ["Launch", "Prep"] or is_workday(d):
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
            else:
                cells.append(f'<td class="{base_cls}"></td>')
        rows.append("<tr>" + "".join(cells) + "</tr>")

    return f"""
    <div class="legend">
      <span class="legend-item"><span class="legend-dot" style="background:{UI_AD2};"></span>Ad2</span>
      <span class="legend-item"><span class="legend-dot" style="background:{UI_CLIENT};"></span>客戶</span>
      <span class="legend-item"><span class="legend-dot" style="background:{UI_LAUNCH};"></span>上線</span>
      <span class="legend-item"><span class="legend-dot" style="background:{UI_PREP};"></span>預備上線</span>
    </div>
    <div class="timeline-wrap">
      <table class="timeline-table">
        {''.join(rows)}
      </table>
    </div>
    """

def sync_task_field(task_id: str, field: str, widget_key: str):
    for row in st.session_state.tasks:
        if row.get("id") == task_id:
            row[field] = st.session_state.get(widget_key)
            break

def add_task():
    st.session_state.tasks.append({
        "id": f"task_new_{uuid.uuid4().hex[:6]}",
        "顯示": True,
        "任務名稱": "",
        "Action By": "Ad2",
        "工作天數": 1,
        "上線日": False,
    })

def move_task_up(idx: int):
    if 0 < idx < len(st.session_state.tasks):
        st.session_state.tasks[idx - 1], st.session_state.tasks[idx] = st.session_state.tasks[idx], st.session_state.tasks[idx - 1]

def move_task_down(idx: int):
    if 0 <= idx < len(st.session_state.tasks) - 1:
        st.session_state.tasks[idx + 1], st.session_state.tasks[idx] = st.session_state.tasks[idx], st.session_state.tasks[idx + 1]

def remove_task(idx: int):
    if 0 <= idx < len(st.session_state.tasks):
        st.session_state.tasks.pop(idx)

def copy_task(idx: int):
    if 0 <= idx < len(st.session_state.tasks):
        row = st.session_state.tasks[idx].copy()
        row["id"] = f"task_copy_{uuid.uuid4().hex[:6]}"
        st.session_state.tasks.insert(idx + 1, row)

def generate_schedule():
    had_previous_output = st.session_state.schedule_df is not None
    holidays = parse_holidays(st.session_state.holidays_text)
    tasks = get_active_tasks()
    calculation_mode = MODE_MAP[st.session_state.mode_display]
    start_date_obj = None if st.session_state.mode_display == "上線日回推" else st.session_state.start_date_value
    launch_date_obj = None if st.session_state.mode_display == "製作日推進" else st.session_state.launch_date_value

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
    st.session_state.display_columns = display_columns
    st.session_state.holidays_dt = holidays_dt
    st.session_state.last_generated_name = st.session_state.project_name or "未命名專案"
    st.session_state.status_msg = "時程表已更新。" if had_previous_output else "已產出時程表。"

def reset_defaults():
    st.session_state.project_name = ""
    st.session_state.mode_display = MODE_OPTIONS[0]
    st.session_state.start_date_value = date.today()
    st.session_state.launch_date_value = date.today()
    st.session_state.collapse_threshold = 2
    st.session_state.tasks = [row.copy() for row in DEFAULT_TASKS]
    st.session_state.schedule_df = None
    st.session_state.excel_bytes = None
    st.session_state.display_columns = None
    st.session_state.holidays_dt = None
    st.session_state.warning_msg = ""
    st.session_state.last_generated_name = "未命名專案"
    st.session_state.status_msg = ""

# =========================
# UI
# =========================
st.title("製作時程排程工具")
st.caption("快速設定專案日期與流程後，即可產出 Excel 時程表。")

with st.sidebar:
    st.subheader("假日設定")
    st.text_area("假日清單（每行一筆，格式：YYYY-MM-DD,名稱）", key="holidays_text", height=420)

with st.container(border=True):
    c1, c2 = st.columns([6,0.65], vertical_alignment="center")
    with c1:
        st.markdown('<div class="section-title">專案設定</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">先決定排程方式與日期，再按下「產出時程表」。</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="mini-reset">', unsafe_allow_html=True)
        st.button("重設", on_click=reset_defaults, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    r1c1, r1c2, r1c3 = st.columns([2.6,1.5,0.9], vertical_alignment="bottom")
    with r1c1:
        st.text_input("專案名稱", key="project_name", placeholder="請輸入專案名稱")
    with r1c2:
        st.selectbox("排程方式", MODE_OPTIONS, key="mode_display")
    with r1c3:
        st.number_input("日期縮略門檻", min_value=1, max_value=30, step=1, key="collapse_threshold")

    start_disabled = st.session_state.mode_display == "上線日回推"
    launch_disabled = st.session_state.mode_display == "製作日推進"

    r2c1, r2c2, r2c3 = st.columns([1.5,1.5,1.0], vertical_alignment="bottom")
    with r2c1:
        st.date_input("開始日期", key="start_date_value", disabled=start_disabled)
    with r2c2:
        st.date_input("上線日期", key="launch_date_value", disabled=launch_disabled)
    with r2c3:
        st.button("產出時程表", type="primary", use_container_width=True, on_click=generate_schedule)

st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)

if st.session_state.warning_msg:
    st.warning(st.session_state.warning_msg)

if st.session_state.status_msg:
    st.success(st.session_state.status_msg)

if st.session_state.schedule_df is not None:
    with st.container(border=True):
        p1, p2 = st.columns([5.4,1.05], vertical_alignment="center")
        with p1:
            st.markdown('<div class="section-title">排程預覽</div>', unsafe_allow_html=True)
        with p2:
            filename = f"{datetime.now().strftime('%m%d')}_{st.session_state.last_generated_name}.xlsx"
            st.download_button("下載 Excel", data=st.session_state.excel_bytes, file_name=filename,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, type="primary")
        st.markdown(
            render_stable_preview(
                st.session_state.schedule_df,
                st.session_state.display_columns,
                st.session_state.holidays_dt,
            ),
            unsafe_allow_html=True,
        )

st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)




with st.container(border=True):
    h1, h2 = st.columns([5,1.05], vertical_alignment="center")
    with h1:
        st.markdown('<div class="section-title">流程設定</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-sub">可新增、複製、刪除、排序與修改任務。</div>', unsafe_allow_html=True)
    with h2:
        st.button("新增任務", on_click=add_task, use_container_width=True)

    st.markdown('<div class="task-config-area flow-config-scope">', unsafe_allow_html=True)
    st.markdown('<div class="task-config-head">', unsafe_allow_html=True)
    hc1, hc2, hc3, hc4, hc5, hc6, hc7, hc8 = st.columns([0.72, 2.9, 1.2, 0.9, 0.9, 1.0, 0.55, 0.45], vertical_alignment="center")
    headers = [
        (hc1, "顯示"),
        (hc2, "任務名稱"),
        (hc3, "Action By"),
        (hc4, "工作天數"),
        (hc5, "上線日"),
        (hc6, "排序"),
        (hc7, "複製"),
        (hc8, "刪除"),
    ]
    for col, label in headers:
        with col:
            st.markdown(f'<div class="task-head-label">{label}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    for idx, row in enumerate(st.session_state.tasks):
        rid = row["id"]
        st.markdown('<div class="task-table-row">', unsafe_allow_html=True)

        c1, c2, c3, c4, c5, c6, c7, c8 = st.columns([0.62, 3.15, 1.2, 0.78, 0.7, 1.0, 0.55, 0.45], vertical_alignment="center")

        with c1:
            key = f"show_{rid}"
            if key not in st.session_state:
                st.session_state[key] = row["顯示"]
            st.checkbox("顯示", key=key, label_visibility="collapsed",
                        on_change=sync_task_field, args=(rid, "顯示", key))

        with c2:
            key = f"task_{rid}"
            if key not in st.session_state:
                st.session_state[key] = row["任務名稱"]
            st.text_input("任務名稱", key=key, label_visibility="collapsed",
                          on_change=sync_task_field, args=(rid, "任務名稱", key))

        with c3:
            key = f"owner_{rid}"
            if key not in st.session_state:
                st.session_state[key] = row["Action By"]
            st.selectbox("Action By", ["Ad2", "客戶"], key=key, label_visibility="collapsed",
                         on_change=sync_task_field, args=(rid, "Action By", key))

        with c4:
            key = f"days_{rid}"
            if key not in st.session_state:
                st.session_state[key] = int(row["工作天數"])
            st.number_input("工作天數", min_value=1, step=1, key=key, label_visibility="collapsed",
                            on_change=sync_task_field, args=(rid, "工作天數", key))

        with c5:
            key = f"launch_{rid}"
            if key not in st.session_state:
                st.session_state[key] = row["上線日"]
            st.checkbox("上線日", key=key, label_visibility="collapsed",
                        on_change=sync_task_field, args=(rid, "上線日", key))

        with c6:
            s1, s2 = st.columns([1, 1], vertical_alignment="center")
            with s1:
                st.markdown('<div class="op-btn">', unsafe_allow_html=True)
                if st.button("↑", key=f"up_{rid}", use_container_width=True, disabled=(idx == 0)):
                    move_task_up(idx)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
            with s2:
                st.markdown('<div class="op-btn">', unsafe_allow_html=True)
                if st.button("↓", key=f"down_{rid}", use_container_width=True, disabled=(idx == len(st.session_state.tasks) - 1)):
                    move_task_down(idx)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

        with c7:
            st.markdown('<div class="op-btn">', unsafe_allow_html=True)
            if st.button("⧉", key=f"copy_{rid}", use_container_width=True):
                copy_task(idx)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        with c8:
            st.markdown('<div class="op-btn">', unsafe_allow_html=True)
            if st.button("✕", key=f"del_{rid}", use_container_width=True):
                remove_task(idx)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)


st.markdown('</div>', unsafe_allow_html=True)
