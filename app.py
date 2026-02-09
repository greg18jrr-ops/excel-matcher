import streamlit as st
import pandas as pd
import io

def process_excel(file):
    try:
        # Load the Excel file
        xls = pd.ExcelFile(file)
        
        # Check if required sheets exist
        sheet_names = xls.sheet_names
        if len(sheet_names) < 2:
            st.error("錯誤：上傳的檔案必須至少包含兩個工作表 (Sheet1 和 Sheet2)。")
            return None

        # Load Sheet1 and Sheet2
        # Assuming first sheet is Sheet1 and second is Sheet2 based on user description, 
        # but user said "工作表1" and "工作表2". Let's try to load by name first, then fall back to index.
        
        sheet1_name = sheet_names[0]
        sheet2_name = sheet_names[1]
        
        # If user specifically named them "工作表1" etc, we should probably check, 
        # but usually index 0 and index 1 is safer for general use unless specified.
        # Let's use index 0 as target (to be updated) and index 1 as source (reference).
        
        df1 = pd.read_excel(file, sheet_name=0)
        df2 = pd.read_excel(file, sheet_name=1)

        # Convert columns to string to ensure matching works correctly
        # Sheet1 Column I is index 8 (0-based)
        # Sheet2 Column A is index 0 (0-based)
        
        # Helper to get column name by index safely
        col_I_name = df1.columns[8] 
        col_A_name = df2.columns[0]

        # Create a dictionary for faster lookup from Sheet2
        # Key: Value in Col A, Value: Row data from Sheet2
        ref_dict = df2.set_index(col_A_name).to_dict('index')

        # Columns to copy from Sheet2 (D, E, F, G, H, I, J) -> Indices 3, 4, 5, 6, 7, 8, 9
        # Columns to paste into Sheet1 (Q, R, S, T, U, V, W) -> Indices 16, 17, 18, 19, 20, 21, 22
        
        source_cols_indices = [3, 4, 5, 6, 7, 8, 9] 
        target_cols_indices = [16, 17, 18, 19, 20, 21, 22]

        # Ensure Sheet1 has enough columns. If not, add them.
        while len(df1.columns) <= max(target_cols_indices):
            df1[f'NewCol_{len(df1.columns)}'] = None

        # Iterate through Sheet1 and update
        for idx, row in df1.iterrows():
            match_val = row[col_I_name]
            
            if match_val in ref_dict:
                source_row = ref_dict[match_val]
                
                # Get source column names to access data from dict
                source_col_names = [df2.columns[i] for i in source_cols_indices]
                
                # Update Sheet1 specific cells
                for i, source_col in enumerate(source_col_names):
                    target_col_idx = target_cols_indices[i]
                    target_col_name = df1.columns[target_col_idx]
                    
                    # Write value
                    df1.at[idx, target_col_name] = source_row[source_col]

        # Save to buffer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Write the updated Sheet1
            df1.to_excel(writer, sheet_name=sheet1_name, index=False)
            # Write the original Sheet2
            df2.to_excel(writer, sheet_name=sheet2_name, index=False)
        
        output.seek(0)
        return output

    except Exception as e:
        st.error(f"發生錯誤: {e}")
        return None

st.title("Excel 資料比對與填入工具 😎")
st.markdown("""
這是一個簡單的工具，功能如下：
1. 上傳 Excel 檔案 (.xlsx)
2. 程式會讀取 **第二個工作表 (Sheet2)** 的 **A欄**
3. 在 **第一個工作表 (Sheet1)** 的 **I欄** 尋找相同的值
4. 若找到，將 Sheet2 的 **D~J欄** 資料填入 Sheet1 的 **Q~W欄**
5. 最後產生包含 **更新後的 Sheet1** 與 **原始 Sheet2** 的合併檔案
""")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file is not None:
    if st.button("開始處理"):
        with st.spinner('處理中...'):
            result = process_excel(uploaded_file)
            
        if result:
            st.success("處理完成！")
            st.download_button(
                label="下載處理後的 Excel",
                data=result,
                file_name="processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
