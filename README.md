# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 重寫預覽表格結構，改成單一 table，不再使用 thead/tbody 分段
- 讓 `～` 的整欄合併在預覽中真正生效
- Excel 輸出維持前一版正確結果
