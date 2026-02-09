# Excel Matcher Tool

這是一個使用 Python Streamlit 建立的簡單 Excel 資料處理工具。

## 功能

1. 上傳 Excel 檔案 (.xlsx)。
2. 系統自動讀取：
   - **工作表 2 (來源)**
   - **工作表 1 (目標)**
3. 比對邏輯：
   - 取 **工作表 2** 的 **A欄** 值。
   - 在 **工作表 1** 的 **I欄** 尋找相同值。
4. 資料填寫：
   - 若比對成功，將 **工作表 2** 的 **D, E, F, G, H, I, J** 欄位數值。
   - 填入 **工作表 1** 對應的 **Q, R, S, T, U, V, W** 欄位。
5. 產出新的 Excel 檔案供下載。

## 如何執行

1. 安裝依賴套件：
   ```bash
   pip install -r requirements.txt
   ```

2. 執行程式：
   ```bash
   streamlit run app.py
   ```
