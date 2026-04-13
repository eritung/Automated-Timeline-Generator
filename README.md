
# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 任務列新增上移／下移，方便快速調整順序
- 預覽假日色塊邏輯比照 Excel，不再把一般任務壓在假日上
- 再次產出時會顯示「時程表已更新」
- 調整欄位與按鈕對齊，重設按鈕縮小
