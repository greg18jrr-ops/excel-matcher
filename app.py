import streamlit as st
import pandas as pd
import io

def process_excel(file):
    try:
        # Load the Excel file
        xls = pd.ExcelFile(file)
        
        # Check if required sheets exist
        sheet_names = xls.sheet_names
        
        # Allow user to select sheets
        st.write("請確認您的工作表對應：")
        
        # Default selections
        default_target_idx = 0
        default_source_idx = 1
        
        if len(sheet_names) > 1:
             for i, name in enumerate(sheet_names):
                 if "1" in name or "一" in name:
                     default_target_idx = i
                     break
             for i, name in enumerate(sheet_names):
                 if "2" in name or "二" in name:
                     default_source_idx = i
                     break

        if default_target_idx == default_source_idx and len(sheet_names) > 1:
             default_source_idx = (default_target_idx + 1) % len(sheet_names)

        target_sheet_name = st.selectbox(
            "請選擇要填寫資料的目標工作表 (通常是 Sheet1)",
            sheet_names,
            index=default_target_idx
        )
        
        source_sheet_name = st.selectbox(
            "請選擇提供資料的來源工作表 (通常是 Sheet2)",
            sheet_names,
            index=default_source_idx
        )
        
        if st.button("開始處理"):
            with st.spinner('處理中...'):
                # Load selected sheets
                df1 = pd.read_excel(file, sheet_name=target_sheet_name)
                df2 = pd.read_excel(file, sheet_name=source_sheet_name)

                # Check column bounds
                if len(df1.columns) <= 8:
                     st.error(f"錯誤：目標工作表 '{target_sheet_name}' 欄位不足，找不到第 I 欄 (第 9 欄)。")
                     return
                if len(df2.columns) <= 0:
                     st.error(f"錯誤：來源工作表 '{source_sheet_name}' 欄位不足，找不到第 A 欄 (第 1 欄)。")
                     return

                # Helper to get column name by index safely
                col_I_name = df1.columns[8] 
                col_A_name = df2.columns[0]

                # Create a dictionary for faster lookup from Sheet2
                ref_dict = df2.set_index(col_A_name).to_dict('index')

                # Columns to copy from Sheet2 (D-J -> indices 3-9)
                source_cols_indices = [3, 4, 5, 6, 7, 8, 9] 
                # Columns to paste into Sheet1 (Q-W -> indices 16-22)
                target_cols_indices = [16, 17, 18, 19, 20, 21, 22]

                # Ensure Sheet1 has enough columns
                while len(df1.columns) <= max(target_cols_indices):
                    df1[f'NewCol_{len(df1.columns)}'] = None

                # Iterate through Sheet1 and update
                for idx, row in df1.iterrows():
                    match_val = row[col_I_name]
                    
                    if match_val in ref_dict:
                        source_row = ref_dict[match_val]
                        source_col_names = [df2.columns[i] for i in source_cols_indices]
                        
                        for i, source_col in enumerate(source_col_names):
                            target_col_idx = target_cols_indices[i]
                            target_col_name = df1.columns[target_col_idx]
                            df1.at[idx, target_col_name] = source_row[source_col]

                # Save to buffer
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df1.to_excel(writer, sheet_name=target_sheet_name, index=False)
                    df2.to_excel(writer, sheet_name=source_sheet_name, index=False)
                
                output.seek(0)
                
                st.success("處理完成！")
                st.download_button(
                    label="下載處理後的 Excel",
                    data=output,
                    file_name="processed_file.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        import traceback
        st.error(f"發生錯誤: {e}")
        st.code(traceback.format_exc()) # Show detailed error for debugging

st.title("Excel 資料比對與填入工具 😎")
st.markdown("""
這是一個簡單的工具，功能如下：
1. 上傳 Excel 檔案 (.xlsx)
2. **請確認工作表對應**：選擇目標工作表 (要被填入的) 與 來源工作表 (提供資料的)。
3. 程式會讀取 **來源工作表** 的 **A欄**
4. 在 **目標工作表** 的 **I欄** 尋找相同的值
5. 若找到，將 來源工作表 的 **D~J欄** 資料填入 目標工作表 的 **Q~W欄**
6. 最後產生包含 **更新後的目標工作表** 與 **原始來源工作表** 的合併檔案
""")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file is not None:
    process_excel(uploaded_file)
