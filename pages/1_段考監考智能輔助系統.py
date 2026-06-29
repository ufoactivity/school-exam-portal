import streamlit as st
import pandas as pd
import numpy as np
import io
import pulp
import traceback
import random
import openpyxl
import re
from datetime import datetime

# ==========================================
# 📌 階段一所需套件 (python-docx)
# ==========================================
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==========================================
# 1. 網頁頁面配置與全局狀態記憶體初始化
# ==========================================
st.set_page_config(page_title="段考試務全能系統", page_icon="🏫", layout="wide")
st.title("🏫 教務處試務組 - 段考試務全能系統 (旗艦整合版)")
st.info("💡 系統已升級為「兩階段作業模式」。請利用下方頁籤切換【階段一：試卷催繳】與【階段二：監考排定】。兩階段獨立運作，互不干擾！115.06.06增修")

# --- 狀態記憶體初始化 (Session State) ---
if 'results_p2' not in st.session_state:
    st.session_state['results_p2'] = None
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0
if 'docx_data_p1' not in st.session_state:
    st.session_state['docx_data_p1'] = None
if 'processed_p1' not in st.session_state:
    st.session_state['processed_p1'] = False

# [新增] 階段二所需之專屬記憶體 (兼課與綁定名單)
if 'time_rules' not in st.session_state:
    st.session_state.time_rules = pd.DataFrame([{"老師": None, "允許日期": "無限制", "允許節次": ""} for _ in range(3)])
if 'bind_rules' not in st.session_state:
    st.session_state.bind_rules = pd.DataFrame([{"老師": None, "班級": None}] * 3)
if 'last_bind_file' not in st.session_state:
    st.session_state.last_bind_file = None

# ==========================================
# 2. 輔助功能定義 (包含最新修復)
# ==========================================
def to_excel_bytes(df, header_df=None):
    output = io.BytesIO()
    if header_df is not None:
        # 將標題列(header_df)與內容(df，現已包含頂部檢核區與名單)合併
        final_out = pd.concat([header_df, df], ignore_index=True)
    else:
        final_out = df
    
    final_out = final_out.fillna("")
    
    # 【防呆修復】：將所有開頭為 "=" 或 "-" 的字串前面補一個空白，徹底避免 Excel 誤認儲存格為「公式」而產生檔案損毀警告
    for col in final_out.columns:
        final_out[col] = final_out[col].apply(lambda x: f" {x}" if isinstance(x, str) and (x.startswith("=") or x.startswith("-")) else x)
