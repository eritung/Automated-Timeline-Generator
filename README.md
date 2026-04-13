# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 流程設定每列改用固定 task id 綁定元件，不再因排序後被舊位置狀態蓋回
- 表頭與欄位使用同一組欄寬，對齊再次修正
- 排序按鈕改為對 task id 對應列生效
- 其餘排程與 Excel 輸出邏輯維持不變
