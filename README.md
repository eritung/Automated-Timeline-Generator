# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- `～` 改為從月份列開始，一直到日期、星期與任務區整欄合併成同一大格
- 預覽與 Excel 的 BREAK 呈現方式一致
- 保留先前的跨月月份標示修正
