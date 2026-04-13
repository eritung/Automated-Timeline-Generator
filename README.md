# 製作時程排程工具（Streamlit）

## 啟動方式
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 這版更新
- 流程設定表頭改為置中
- 流程設定列加入斑馬紋底色，方便查看
- 新增複製列功能，可快速複製一筆任務到下一列
- 排序與既有排程邏輯維持不變
