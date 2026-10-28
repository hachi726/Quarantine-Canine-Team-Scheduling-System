# app.py
from datetime import datetime
import io
from pathlib import Path
import sys
import streamlit as st

# 匯入你的核心排班模組
# import solver_core 

st.set_page_config(page_title="智慧排班系統", page_icon="📅", layout="wide")

st.title("📅 智慧自動排班系統（網頁版）")
st.write("上傳你的排班輸入 Excel 模板，系統將自動進行約束求解、衝突診斷並產出最佳班表！")

uploaded_file = st.file_uploader("請上傳排班輸入模板 (Excel 檔)", type=["xlsx"])

if uploaded_file is not None:
    file_path = Path("roster_input_template.xlsx")
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    st.success("✅ 檔案上傳成功！")

    if st.button("🚀 開始執行自動排班與求解", type="primary"):
        # 建立一個容器用來顯示執行過程中的 print 訊息
        log_container = st.empty()
        output_buffer = io.StringIO()

        # 攔截原本程式裡的 print 輸出，讓它顯示在網頁上
        sys.stdout = output_buffer
        
        try:
            # 呼叫你封裝好的核心排班邏輯
            # success, message = solver_core.run_solver(file_path)
            
            # 這裡模擬執行成功
            success = True
            print("開始排班：2026 年 11 月，共 21 人。")
            print("求解器正在尋找最佳可行班表...")
            print("排班完成！已順利產出可行班表。")
            
        except Exception as e:
            success = False
            print(f"❌ 發生錯誤：{e}")
            
        # 恢復標準輸出
        sys.stdout = sys.__stdout__

        # 顯示日誌
        log_container.text_area("執行日誌與衝突診斷報告", output_buffer.getvalue(), height=300)

        if success and file_path.exists():
            st.balloons()
            st.success("🎉 排班完成！你可以點擊下方按鈕下載最新的 Excel 班表。")
            
            with open(file_path, "rb") as f:
                st.download_button(
                    label="📥 下載產出後的排班結果 Excel",
                    data=f,
                    file_name=f"roster_result_{datetime.now().strftime('%Y%m')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("💡 請先上傳你的 Excel 模板檔案以開始操作。")