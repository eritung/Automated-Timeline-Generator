# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 修正 Excel 跨月時第二個月份標示可能消失的問題
- 將月份 merge 邏輯改為先切出 month segments 再輸出
- 預覽中的 BREAK 欄改成整段垂直合併的「～」，比照 Excel 輸出效果
