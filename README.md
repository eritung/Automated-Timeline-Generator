# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 拿掉自訂 HTML 卡片包覆，改用 Streamlit 原生 container(border=True)
- 修正每個區塊頂部出現白色長條的問題
- Excel 匯出的上線色塊維持原本純紅
