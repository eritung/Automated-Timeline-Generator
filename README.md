# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 流程設定表頭與欄位改用同一組欄寬，對齊修正
- 排序按鈕改成直接觸發上移／下移
- 重設按鈕縮小
- 其餘排程與 Excel 輸出邏輯維持不變
