import streamlit as st
import pandas as pd
import io
import traceback
import re
from datetime import datetime

# ==========================================
# 1. 網頁頁面配置與記憶體初始化
# ==========================================
st.set_page_config(page_title="模擬考調查智能系統", page_icon="📊", layout="wide")
st.title("📊 教務處-模擬考調查智能輔助系統 (雙工作表預填套印版)")
st.info("💡 試務組終極進化：支援上傳「雙工作表設定檔」(工作表1:科別對應, 工作表2:費用對應)，系統會自動跨表比對並套印至調查表中！")

# --- 初始化系統記憶體 (防重整閃退) ---
if 'mock_processed' not in st.session_state:
    st.session_state.mock_processed = False
if 'mock_excel_data' not in st.session_state:
    st.session_state.mock_excel_data = None
if 'mock_uploader_key' not in st.session_state:
    st.session_state.mock_uploader_key = 0
if 'mock_preview_df' not in st.session_state:
    st.session_state.mock_preview_df = None

# 階段一專用記憶體
if 'template_processed' not in st.session_state:
    st.session_state.template_processed = False
if 'template_excel_data' not in st.session_state:
    st.session_state.template_excel_data = None

# ==========================================
# 2. 輔助功能定義 (防呆與髒數據處理)
# ==========================================
def get_str_col(df, keywords):
    """模糊比對欄位名稱，防禦老師上傳不同格式的表單"""
    if isinstance(keywords, str): keywords = [keywords]
    for kw in keywords:
        for i, col in enumerate(df.columns):
            if kw == str(col).strip():
                return df.iloc[:, i].fillna("").astype(str).str.strip()
    for kw in keywords:
        for i, col in enumerate(df.columns):
            if kw in str(col):
                return df.iloc[:, i].fillna("").astype(str).str.strip()
    return pd.Series([""] * len(df), index=df.index)

def clean_class_name(c):
    if pd.isna(c): return ""
    s = str(c).strip().replace('ㄧ', '一').replace(' ', '').replace('　', '')
    s = s.translate(str.maketrans('１２３４５６７８９０', '1234567890'))
    return s

# 統測標準20群類 + 51~56 代碼表
STANDARD_MAPPING_VOC = {
    '1': '機械群', '2': '動力機械群', '3': '電機與電子群(電機類)', '4': '電機與電子群(資電類)',
    '5': '化工群', '6': '土木與建築群', '7': '設計群', '8': '工程與管理類',
    '9': '商業與管理群', '10': '衛生與護理類', '11': '食品群', '12': '家政群幼保類',
    '13': '家政群生活應用類', '14': '農業群', '15': '外語群(英文類)', '16': '外語群(日文類)',
    '17': '餐旅群', '18': '海事群', '19': '水產群', '20': '藝術群(影視類)',
    '51': '電機與電子群(03+04類)', '52': '家政群(12+13類)', '53': '商管外語群(1)-(9+15類)',
    '54': '商管外語群(2)-(9+16類)', '55': '商管外語群(3)-(15+16類)', '56': '商管外語群(4)-(9+15+16類)'
}

# 學測常見考科組合代碼表
STANDARD_MAPPING_GEN = {
    '1': '社會組(國英數B社)',
    '2': '自然組(國英數A自)',
    '3': '跨考組(國英數A社)',
    '4': '跨考組(國英數B自)',
    '5': '全考(國英數A數B社自)'
}

# ==========================================
# 3. 雙階段頁籤設計
# ==========================================
tab1, tab2 = st.tabs(["📄 階段一：從名條產出【空白/預填意願調查表】", "💰 階段二：從調查表產出【試務與收費報表】"])

# ---------------------------------------------------------
# 【階段一：產出調查表】
# ---------------------------------------------------------
with tab1:
    st.subheader("🛠️ 製作公版模擬考意願調查表 (支援跨表自動套印)")
    st.markdown("上傳學生名單與「預設對照表（工作表1=科別預設類組、工作表2=類組單次費用）」，系統將為您自動排版並預填好各班學生的類組與費用！")
    
    col1_t1, col2_t1 = st.columns([1, 1], gap="large")
    
    with col1_t1:
        file_roster = st.file_uploader("📥 1. 上傳學生原始名條 (必填)", type=['xlsx', 'csv'], key="roster_uploader")
        file_preset = st.file_uploader("📥 2. 上傳【雙工作表預設對照檔】 (選填)", type=['xlsx'], key="preset_uploader")
        school_type = st.radio("🏫 選擇產出的學制類型：", ["技高 (統測群類)", "普高 (學測考科)"], horizontal=True)
        
    with col2_t1:
        default_title = "114學年度國立華南高商統測模擬考 報考類組調查表" if "技高" in school_type else "114學年度國立華南高商學測模擬考 報考考科調查表"
        template_name = st.text_input("🎯 擬定表單大標題", value=default_title)
        default_price = st.number_input("💰 預設單次費用 (無對照檔或查無資料時套用，可留 0)", min_value=0, max_value=2000, value=0, step=10)
        
    if st.button("🚀 生成一班一頁調查表", type="primary", use_container_width=True, key="btn_gen_template"):
        if not file_roster:
            st.error("🚨 請先上傳學生原始名條！")
        else
