from datetime import datetime
import io
from pathlib import Path
import sys
import streamlit as st
import solver_core

st.set_page_config(page_title="智慧排班系統", page_icon="📅", layout="wide")

st.title("📅 智慧自動排班系統（網頁版）")
st.write("上傳你的排班輸入模板與上個月結果，系統將自動進行跨月銜接與排班求解！")

st.subheader("📁 檔案上傳區塊")
current_file = st.file_uploader("1. 請上傳【當月】排班輸入模板 (Excel 檔)", type=["xlsx"], key="current")
previous_file = st.file_uploader("2. 【選填】請上傳【上個月】的排班結果 Excel（用於自動跨月銜接）", type=["xlsx"], key="previous")

if current_file is not None:
    current_path = Path("roster_input_template.xlsx")
    with open(current_path, "wb") as f:
        f.write(current_file.getbuffer())
        
    previous_path = None
    if previous_file is not None:
        previous_path = Path("previous_roster_result.xlsx")
        with open(previous_path, "wb") as f:
            f.write(previous_file.getbuffer())
    
    st.success("✅ 檔案上傳成功！")

    if st.button("🚀 開始執行自動排班與求解", type="primary"):
        log_container = st.empty()
        output_buffer = io.StringIO()

        sys.stdout = output_buffer
        try:
            solver_core.run_solver(current_path, previous_path)
            success = True
        except Exception as e:
            success = False
            print(f"❌ 發生錯誤：{e}")
            
        sys.stdout = sys.__stdout__

        log_container.text_area("執行日誌與衝突診斷報告", output_buffer.getvalue(), height=300)

        if success and current_path.exists():
            st.balloons()
            st.success("🎉 排班完成！")
            
            with open(current_path, "rb") as f:
                st.download_button(
                    label="📥 下載產出後的排班結果 Excel",
                    data=f,
                    file_name=f"roster_result_{datetime.now().strftime('%Y%m')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("💡 請先上傳當月的 Excel 模板以開始操作。")