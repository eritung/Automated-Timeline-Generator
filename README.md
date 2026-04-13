# 製作時程排程工具（Streamlit）

## 啟動方式

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版調整
- 全繁體中文介面
- 將原本的「類型」改為「上線日（是的話就勾選）」
- 將「產出時程表」按鈕移到畫面上方
- 移除不必要的啟用勾選
- 暫時不保留匯入設定 JSON，讓操作更單純
