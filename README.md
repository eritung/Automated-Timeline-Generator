# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 修正預覽在「同時指定開始與上線日期」模式下，BREAK 欄造成後段日期錯位的問題
- 預覽的 BREAK 欄改成任務區整段垂直合併的「～」
- 預覽月份列不再在 BREAK 位置顯示「～」，後方月份會照正常位置接續
- Excel 跨月月份標示修正仍保留
