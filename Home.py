import streamlit as st

# ==========================================
# 1. 網頁頁面配置 (極簡現代風科技感)
# ==========================================
st.set_page_config(
    page_title="教務處試務組-智慧戰情門戶",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 使用內嵌 CSS 打造現代科技感漸層、極簡字體體系與放大字體
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    
    /* --- 主標題與副標題 --- */
    .main-title {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
        font-size: 3.8rem !important;  /* 🌟 把主標題字體放大了 (原本是 2.8rem) */
        font-weight: 700 !important;
        background: linear-gradient(135deg, #1565c0 0%, #00bcd4 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem !important;
        letter-spacing: -0.05rem;
    }
    .sub-title {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
        font-size: 1.1rem !important;  /* 🌟 英文副標題也稍微等比例放大 */
        font-weight: 600;
        color: #888888;
        text-transform: uppercase;
        letter-spacing: 0.15rem;
        margin-bottom: 2.5rem !important;
    }
    .module-desc {
        color: #666666;
        font-size: 0.95rem;
        line-height: 1.6;
    }

    /* --- 左側選單 (Sidebar) 字體放大 --- */
    [data-testid="stSidebarNav"] span {
        font-size: 1.15rem !important;  
        font-weight: 600 !important;    
        line-height: 2 !important;      
    }
    
    [data-testid="stSidebar"] {
        font-size: 1.1rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 歡迎與標頭區塊
# ==========================================
st.markdown('<h1 class="main-title">⚡ 試務組 AI 智能輔助平台</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Intelligent Examination Administration Ecosystem</p>', unsafe_allow_html=True)
st.divider()

# ==========================================
# 3. 核心工具矩陣入口 (單一入口網導覽)
# ==========================================
st.markdown("### ⚙️ 核心自動化模組 / Production Modules")
st.write("")

col1, col2 = st.columns(2, gap="large")

with col1:
    with st.container(border=True):
        st.markdown("### 📅 段考監考智慧輔助系統")
        st.markdown("""
        <p class="module-desc">
        整合 AI 線性規劃算法，全自動最佳化全校排班與座位分發。
        </p>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        - 🧠 **AI 最佳化對位**：自動平衡全校教師堂數配額與連堂規則。
        - 🖨️ **原格式套印**：公布總表、一覽表 100% 完美保留原 Excel 框線。
        - 🏷️ **標籤全聯動**：交叉比對配課表，一鍵合成試卷袋列印貼紙。
        """)
        st.write("")
        st.page_link("pages/1_監考智能輔助系統.py", label="啟動段考排班作業 →", icon="🚀")

with col2:
    with st.container(border=True):
        st.markdown("### 📝 補考作業智慧輔助系統")
        st.markdown("""
        <p class="module-desc">
        一鍵清洗多科補考學生複雜數據，智慧重組考場與座位配置。
        </p>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        - 🧼 **智慧數據清洗**：自動剃除幽靈名單與全半形錯字干擾。
        - 🏫 **多重考場分流**：自動計算最大考場容量，防範座位重疊衝突。
        - 🖨️ **考務報表輸出**：自動生成個人補考通知單與考場對照表。
        """)
        st.write("")
        st.page_link("pages/2_補考智能輔助系統.py", label="啟動補考處理作業 →", icon="⚡")

# ==========================================
# 4. 底部沉穩頁尾
# ==========================================
st.divider()
st.markdown(
    """
    <div style="text-align: center; color: #bbb; font-size: 0.8rem; letter-spacing: 0.05rem;">
        © 2026 國立華南高商 教務處試務組 · Data-Driven Examination Administration
    </div>
    """, 
    unsafe_allow_html=True
)
