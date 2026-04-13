
# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 預覽改成穩定版樣式，不再硬追 Excel 的複雜合併表頭
- 流程設定改成表單式編輯，減少複製貼上噴錯與資料跳掉
- 新增批次貼上任務區，可貼入多行任務快速套用
- Excel 輸出邏輯維持不變
