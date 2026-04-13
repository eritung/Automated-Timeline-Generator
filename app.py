import io
import json
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import List, Dict, Optional

import pandas as pd
import streamlit as st
import xlsxwriter

DEFAULT_PROJECT_NAME = "未命名專案"
DEFAULT_COLLAPSE_THRESHOLD = 2

DEFAULT_HOLIDAYS = {
    '2026-01-01': '元旦',
    '2026-02-15': '春節連假', '2026-02-16': '春節連假', '2026-02-17': '春節連假',
    '2026-02-18': '春節連假', '2026-02-19': '春節連假', '2026-02-20': '春節連假',
    '2026-02-27': '二二八連假', '2026-02-28': '二二八連假',
    '2026-04-03': '清明連假', '2026-04-04': '清明連假', '2026-04-05': '清明連假',
    '2026-05-01': '勞動節放假',
    '2026-06-19': '端午節連假',
}

DEFAULT_TASKS = [
    {"task": "提供素材", "owner": "客戶", "days": 1, "enabled": True, "type": "normal"},
    {"task": "視覺製作", "owner": "Ad2", "days": 3, "enabled": True, "type": "normal"},
    {"task": "客戶確認", "owner": "客戶", "days": 1, "enabled": True, "type": "normal"},
    {"task": "視覺調整", "owner": "Ad2", "days": 2, "enabled": True, "type": "normal"},
    {"task": "客戶確認", "owner": "客戶", "days": 1, "enabled": True, "type": "normal"},
    {"task": "廣告進稿", "owner": "Ad2", "days": 1, "enabled": True, "type": "normal"},
    {"task": "廣告上線", "owner": "Ad2", "days": 1, "enabled": True, "type": "launch"},
]

COLORS = {
    "client_bar": '#EA9B56',
    "ad2_bar": '#4BACC6',
    "launch_bar": '#FF0000',
    "prep_bar": '#92D050',
    "weekend": '#D9D9D9',
    "holiday_text": '#595959',
}
MONTH_COLORS = ['#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#E7E6E6']


def to_date(v: Optional[str]) -> Optional[date]:
    if not v:
        return None
    return pd.to_datetime(v).date()


def holidays_to_dates(holidays_config: Dict[str, str]) -> List[date]:
    return [pd.to_datetime(h).date() for h in holidays_config.keys()]


def is_workday(d: date, holidays_dt: List[date]) -> bool:
    return d.weekday() < 5 and d not in holidays_dt


def add_workdays(start_date: date, days: int, holidays_dt: List[date]) -> date:
    current = start_date
    if days <= 1:
        return current
    remaining = days - 1
    while remaining > 0:
        current += timedelta(days=1)
        if is_workday(current, holidays_dt):
            remaining -= 1
    return current


def subtract_workdays(start_date: date, days: int, holidays_dt: List[date]) -> date:
    current = start_date
    if days <= 1:
        return current
    remaining = days - 1
    while remaining > 0:
        current -= timedelta(days=1)
        if is_workday(current, holidays_dt):
            remaining -= 1
    return current


def get_previous_workday(d: date, holidays_dt: List[date]) -> date:
    d -= timedelta(days=1)
    while not is_workday(d, holidays_dt):
        d -= timedelta(days=1)
    return d


def get_next_workday(d: date, holidays_dt: List[date]) -> date:
    d += timedelta(days=1)
    while not is_workday(d, holidays_dt):
        d += timedelta(days=1)
    return d


def ensure_workday_forward(d: date, holidays_dt: List[date]) -> date:
    while not is_workday(d, holidays_dt):
        d += timedelta(days=1)
    return d


def generate_schedule(project_name: str, mode: str, target_start_date: Optional[str], target_launch_date: Optional[str], tasks: List[dict], holidays_config: Dict[str, str]):
    tasks_config = [t for t in tasks if t.get('enabled', True)]
    if not tasks_config:
        raise ValueError('至少要有一筆啟用中的任務。')

    start_date_obj = to_date(target_start_date)
    launch_date_obj = to_date(target_launch_date)
    holidays_dt = holidays_to_dates(holidays_config)

    schedule = []
    warning_msg = ""

    if mode == 'backward':
        if not launch_date_obj:
            raise ValueError('Backward 模式需要上線日。')
        current_end = launch_date_obj
        reversed_tasks = tasks_config[::-1]
        temp_schedule = []
        for i, t in enumerate(reversed_tasks):
            duration = int(t['days'])
            is_launch_task = t.get('type') == 'launch'
            if i == 0:
                end_date = current_end
                start_date = subtract_workdays(end_date, duration, holidays_dt)
            else:
                prev_start = temp_schedule[-1]['Start Date']
                end_date = get_previous_workday(prev_start, holidays_dt)
                start_date = subtract_workdays(end_date, duration, holidays_dt)
            temp_schedule.append({
                'Task': t['task'],
                'Owner': t['owner'],
                'Start Date': start_date,
                'End Date': end_date,
                'Type': 'Launch' if is_launch_task else 'Normal'
            })
        schedule = temp_schedule[::-1]

    elif mode == 'forward':
        curr_start = ensure_workday_forward(start_date_obj or date.today(), holidays_dt)
        prev_end = None
        for idx, t in enumerate(tasks_config):
            is_launch_task = t.get('type') == 'launch'
            if is_launch_task and launch_date_obj:
                start_d = launch_date_obj
            else:
                start_d = curr_start if idx == 0 else get_next_workday(prev_end, holidays_dt)
            end_d = add_workdays(start_d, int(t['days']), holidays_dt)
            schedule.append({
                'Task': t['task'],
                'Owner': t['owner'],
                'Start Date': start_d,
                'End Date': end_d,
                'Type': 'Launch' if is_launch_task else 'Normal'
            })
            prev_end = end_d

    elif mode == 'double':
        if not start_date_obj or not launch_date_obj:
            raise ValueError('Double 模式需要開始日與上線日。')
        curr_start = ensure_workday_forward(start_date_obj, holidays_dt)
        prev_end = None
        launch_index = next((i for i, t in enumerate(tasks_config) if t.get('type') == 'launch'), len(tasks_config)-1)
        normal_tasks = tasks_config[:launch_index]
        launch_task_config = tasks_config[launch_index]

        for idx, t in enumerate(normal_tasks):
            start_d = curr_start if idx == 0 else get_next_workday(prev_end, holidays_dt)
            end_d = add_workdays(start_d, int(t['days']), holidays_dt)
            schedule.append({
                'Task': t['task'],
                'Owner': t['owner'],
                'Start Date': start_d,
                'End Date': end_d,
                'Type': 'Normal'
            })
            prev_end = end_d

        if schedule:
            last_task_end = schedule[-1]['End Date']
            real_prep_start = last_task_end + timedelta(days=1)
        else:
            last_task_end = start_date_obj
            real_prep_start = start_date_obj
        real_prep_end = launch_date_obj - timedelta(days=1)

        if last_task_end >= launch_date_obj:
            overrun_days = (last_task_end - launch_date_obj).days
            warning_msg = f'⚠️【時程衝突警告】工作將進行到 {last_task_end}，比上線日晚了 {overrun_days} 天。'

        if real_prep_end >= real_prep_start:
            schedule.append({
                'Task': '預備上線',
                'Owner': 'Ad2',
                'Start Date': real_prep_start,
                'End Date': real_prep_end,
                'Type': 'Prep'
            })

        launch_end = add_workdays(launch_date_obj, int(launch_task_config['days']), holidays_dt)
        schedule.append({
            'Task': launch_task_config['task'],
            'Owner': launch_task_config['owner'],
            'Start Date': launch_date_obj,
            'End Date': launch_end,
            'Type': 'Launch'
        })
    else:
        raise ValueError('未知模式。')

    return pd.DataFrame(schedule), warning_msg


def build_excel(df_schedule: pd.DataFrame, project_name: str, holidays_config: Dict[str, str], collapse_threshold: int = DEFAULT_COLLAPSE_THRESHOLD) -> bytes:
    holidays_dt = holidays_to_dates(holidays_config)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet('時程表')

    font = 'Microsoft JhengHei'
    font_size = 11
    border_fmt = {'border': 1, 'border_color': '#000000'}

    def F(**kwargs):
        return workbook.add_format({'font_name': font, 'font_size': font_size, **kwargs})

    fmt_center = F(align='center', valign='vcenter', **border_fmt)
    fmt_left = F(align='left', valign='vcenter', **border_fmt)
    fmt_weekend = F(bg_color=COLORS['weekend'], align='center', valign='vcenter', **border_fmt)
    fmt_date_num = F(align='center', valign='vcenter', **border_fmt)
    fmt_holiday_merged = F(align='center', valign='vcenter', text_wrap=True, bg_color=COLORS['weekend'], border=1, font_color=COLORS['holiday_text'], bold=True)
    fmt_header_main = F(bold=True, align='center', valign='vcenter', bg_color='#FFFFFF', **border_fmt)
    fmt_bar_client = F(bg_color=COLORS['client_bar'], align='center', valign='vcenter', **border_fmt)
    fmt_bar_ad2 = F(bg_color=COLORS['ad2_bar'], align='center', valign='vcenter', **border_fmt)
    fmt_bar_launch = F(bg_color=COLORS['launch_bar'], align='center', valign='vcenter', **border_fmt)
    fmt_bar_prep = F(bg_color=COLORS['prep_bar'], align='center', valign='vcenter', **border_fmt)
    fmt_break_merge = F(align='center', valign='vcenter', **border_fmt)
    fmt_legend_client = F(bg_color=COLORS['client_bar'], align='center', valign='vcenter', **border_fmt)
    fmt_legend_ad2 = F(bg_color=COLORS['ad2_bar'], align='center', valign='vcenter', **border_fmt)

    def _is_workday(d: date) -> bool:
        return is_workday(d, holidays_dt)

    min_date = df_schedule['Start Date'].min()
    max_date = df_schedule['End Date'].max()
    full_dates = list(pd.date_range(min_date, max_date, freq='D'))

    schedule = df_schedule.to_dict('records')
    display_columns = []
    prep_task = next((item for item in schedule if item['Type'] == 'Prep'), None)
    if prep_task and (prep_task['End Date'] - prep_task['Start Date']).days + 1 >= collapse_threshold:
        keep_start = prep_task['Start Date']
        resume_date = prep_task['End Date'] + timedelta(days=1)
        break_added = False
        for d in full_dates:
            d_date = d.date()
            if d_date >= resume_date or d_date <= keep_start:
                display_columns.append(d)
            else:
                if not break_added:
                    display_columns.append('BREAK')
                    break_added = True
    else:
        display_columns = full_dates

    worksheet.write(0, 2, '客戶', fmt_legend_client)
    worksheet.write(0, 3, 'Ad2', fmt_legend_ad2)
    worksheet.merge_range(1, 0, 3, 0, '製作時程', fmt_header_main)
    worksheet.merge_range(1, 1, 3, 1, 'Action by', fmt_header_main)
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 12)

    col_start, row_start = 2, 4
    current_month, merge_start_col, month_color_idx = None, col_start, 0
    break_cols_excel = []

    for i, item in enumerate(display_columns):
        col = col_start + i
        if item == 'BREAK':
            worksheet.set_column(col, col, 4)
            break_cols_excel.append(col)
            continue
        d = item
        if current_month is None:
            current_month = d.month
        if d.month != current_month:
            month_fmt = F(bold=True, align='center', valign='vcenter', bg_color=MONTH_COLORS[month_color_idx % len(MONTH_COLORS)], **border_fmt)
            worksheet.merge_range(1, merge_start_col, 1, col - 1, date(d.year, current_month, 1).strftime('%b').upper(), month_fmt)
            current_month, merge_start_col, month_color_idx = d.month, col, month_color_idx + 1
        if i == len(display_columns) - 1:
            month_fmt = F(bold=True, align='center', valign='vcenter', bg_color=MONTH_COLORS[month_color_idx % len(MONTH_COLORS)], **border_fmt)
            worksheet.merge_range(1, merge_start_col, 1, col, d.strftime('%b').upper(), month_fmt)
        is_h = not _is_workday(d.date())
        worksheet.write(2, col, d.day, fmt_weekend if is_h else fmt_date_num)
        worksheet.write(3, col, {0:'一',1:'二',2:'三',3:'四',4:'五',5:'六',6:'日'}[d.weekday()], fmt_weekend if is_h else fmt_date_num)
        worksheet.set_column(col, col, 4.5)

    last_task_row = row_start + len(schedule) - 1
    for c in break_cols_excel:
        worksheet.merge_range(2, c, last_task_row, c, '～', fmt_break_merge)

    for idx, item in enumerate(schedule):
        row = row_start + idx
        worksheet.write(row, 0, item['Task'], fmt_left)
        worksheet.write(row, 1, item['Owner'], fmt_left)
        bar_fmt = fmt_bar_launch if item['Type'] == 'Launch' else fmt_bar_prep if item['Type'] == 'Prep' else fmt_bar_client if '客戶' in item['Owner'] else fmt_bar_ad2
        for i, col_item in enumerate(display_columns):
            if col_item == 'BREAK':
                continue
            col = col_start + i
            d_date = col_item.date()
            if item['Start Date'] <= d_date <= item['End Date']:
                if item['Type'] in ['Launch', 'Prep'] or _is_workday(d_date):
                    worksheet.write(row, col, '', bar_fmt)
                else:
                    worksheet.write(row, col, '', fmt_weekend)
            else:
                worksheet.write(row, col, '', fmt_weekend if not _is_workday(d_date) else fmt_center)

    holiday_blocks_info = []
    current_block_start = -1
    current_block_dates = []
    for i, col_item in enumerate(display_columns):
        is_holiday_day = False
        if col_item != 'BREAK':
            d_date = col_item.date()
            if not _is_workday(d_date):
                is_holiday_day = True
        if is_holiday_day:
            if current_block_start == -1:
                current_block_start = i
            current_block_dates.append(col_item.date())
        else:
            if current_block_start != -1:
                found_name = next((holidays_config[d.strftime('%Y-%m-%d')] for d in current_block_dates if d.strftime('%Y-%m-%d') in holidays_config), None)
                if found_name:
                    holiday_blocks_info.append({'start_col': current_block_start, 'end_col': i - 1, 'name': '\n'.join(list(found_name))})
            current_block_start, current_block_dates = -1, []

    for block in holiday_blocks_info:
        c1, c2 = col_start + block['start_col'], col_start + block['end_col']
        worksheet.merge_range(row_start, c1, last_task_row, c2, block['name'], fmt_holiday_merged)

    writer.close()
    output.seek(0)
    return output.read()


def default_config():
    return {
        'project_name': DEFAULT_PROJECT_NAME,
        'mode': 'forward',
        'target_start_date': str(date.today()),
        'target_launch_date': '',
        'collapse_threshold': DEFAULT_COLLAPSE_THRESHOLD,
        'tasks': DEFAULT_TASKS,
        'holidays': DEFAULT_HOLIDAYS,
    }


st.set_page_config(page_title='時程表產生器', layout='wide')
st.title('時程表產生器')
st.caption('把 Colab 版搬成可直接操作的網頁工具。')

if 'config' not in st.session_state:
    st.session_state.config = default_config()

with st.sidebar:
    st.subheader('設定')
    uploaded = st.file_uploader('匯入設定 JSON', type=['json'])
    if uploaded:
        st.session_state.config = json.load(uploaded)
        st.success('已載入設定。')
    if st.button('重設為預設值'):
        st.session_state.config = default_config()
        st.rerun()

cfg = st.session_state.config

col1, col2, col3 = st.columns(3)
with col1:
    project_name = st.text_input('專案名稱', value=cfg['project_name'])
with col2:
    mode = st.selectbox('模式', ['forward', 'backward', 'double'], index=['forward', 'backward', 'double'].index(cfg['mode']))
with col3:
    collapse_threshold = st.number_input('縮略門檻', min_value=1, value=int(cfg.get('collapse_threshold', 2)))

col4, col5 = st.columns(2)
with col4:
    target_start_date = st.text_input('開始日', value=cfg.get('target_start_date', ''))
with col5:
    target_launch_date = st.text_input('上線日', value=cfg.get('target_launch_date', ''))

st.subheader('任務流程')
tasks_df = pd.DataFrame(cfg['tasks'])
edited_tasks = st.data_editor(
    tasks_df,
    num_rows='dynamic',
    use_container_width=True,
    column_config={
        'task': st.column_config.TextColumn('任務'),
        'owner': st.column_config.TextColumn('負責人'),
        'days': st.column_config.NumberColumn('天數', min_value=1, step=1),
        'enabled': st.column_config.CheckboxColumn('啟用'),
        'type': st.column_config.SelectboxColumn('類型', options=['normal', 'launch'])
    }
)

st.subheader('假日設定')
holidays_df = pd.DataFrame([
    {'date': k, 'name': v} for k, v in cfg['holidays'].items()
])
edited_holidays = st.data_editor(
    holidays_df,
    num_rows='dynamic',
    use_container_width=True,
    column_config={
        'date': st.column_config.TextColumn('日期 YYYY-MM-DD'),
        'name': st.column_config.TextColumn('名稱')
    }
)

run = st.button('產出時程表', type='primary')

if run:
    try:
        holidays_dict = {str(row['date']): str(row['name']) for _, row in edited_holidays.iterrows() if pd.notna(row['date']) and pd.notna(row['name'])}
        tasks_list = edited_tasks.fillna('').to_dict('records')
        df_schedule, warning_msg = generate_schedule(project_name, mode, target_start_date, target_launch_date, tasks_list, holidays_dict)
        st.success('排程完成。')
        if warning_msg:
            st.warning(warning_msg)
        st.dataframe(df_schedule, use_container_width=True)

        excel_bytes = build_excel(df_schedule, project_name, holidays_dict, int(collapse_threshold))
        now_str = (datetime.utcnow() + timedelta(hours=8)).strftime('%m%d')
        st.download_button(
            '下載 Excel',
            data=excel_bytes,
            file_name=f'{now_str}_{project_name}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        export_config = {
            'project_name': project_name,
            'mode': mode,
            'target_start_date': target_start_date,
            'target_launch_date': target_launch_date,
            'collapse_threshold': int(collapse_threshold),
            'tasks': tasks_list,
            'holidays': holidays_dict,
        }
        st.download_button(
            '下載設定 JSON',
            data=json.dumps(export_config, ensure_ascii=False, indent=2),
            file_name=f'{project_name}_config.json',
            mime='application/json'
        )
        st.session_state.config = export_config
    except Exception as e:
        st.error(f'發生錯誤：{e}')
