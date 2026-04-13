# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 流程設定欄位改為：任務名稱 / Action By / 工作天數 / 上線日
- 假日設定移到可收合側邊欄
- 在正推或回推模式下，不需要的日期欄位會鎖定
- 重設按鈕縮小並移到專案設定右上角
- 排程預覽與下載 Excel 按鈕移到專案設定正下方
