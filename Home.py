import streamlit as st
import os
import base64

# ==========================================
# 1. 網頁頁面配置 (極簡現代風科技感)
# ==========================================
st.set_page_config(
    page_title="教務處試務組-智能輔助平台",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 將圖片轉為 Base64 格式，這是讓圖片強制滿版的核心技術 ---
@st.cache_data
def get_image_base64(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return None

# 使用內嵌 CSS 打造現代科技感漸層、極簡字體體系與放大字體
st.markdown("""
    <style>
    /* 載入繁體中文與活潑趣味的快樂體字型 (ZCOOL KuaiLe) */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&family=ZCOOL+KuaiLe&display=swap');
    
    /* --- 減少 Streamlit 預設頂部邊距，讓標題大幅上移 --- */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
    }

    /* --- 全域設定 --- */
    html, body, [data-testid="stSidebar"] {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
    }

    /* --- 主標題與副標題 --- */
    .main-title {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
        font-size: 3.8rem !important;
        font-weight: 700 !important;
        background: linear-gradient(135deg, #1565c0 0%, #00bcd4 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem !important;
        letter-spacing: -0.05rem;
    }
    .sub-title {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
        font-size: 1.1rem !important;
        font-weight: 600;
        color: #888888;
        text-transform: uppercase;
        letter-spacing: 0.15rem;
        margin-bottom: 1.0rem !important; 
    }
    
    /* --- 橫幅容器設定 (用於包裹文字註記) --- */
    .hero-banner-container {
        position: relative;              
        width: 100% !important;
        border-radius: 15px;
        overflow: hidden;                
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin-bottom: 1.5rem;
    }
    
    /* 通過純 HTML Class 設定終極滿版橫幅 */
    .hero-banner {
        width: 100% !important;          
        height: 250px !important;        
        object-fit: cover !important;    
        
        /* 🌟 關鍵修正：畫面移至 100% 完美觸底，階梯與地面景色呈現最完整 🌟 */
        object-position: 50% 100% !important; 
        
        display: block;                  
    }
    
    /* --- 🌟 關鍵修正：優化後的右下角文字層（防黏糊、再往下） 🌟 --- */
    .banner-overlay-text {
        position: absolute;              
        bottom: 12px;                    /* 🌟 關鍵修正：文字再往下移，完美貼近底緣 🌟 */
        right: 25px;                     
        color: white;                    
        
        font-family: 'ZCOOL KuaiLe', 'Microsoft JhengHei', sans-serif !important;
        
        /* 🌟 關鍵修正：調小粗度至 400，並拉大字距至 0.15rem，徹底解決字體黏糊問題 🌟 */
        font-weight: 400 !important;     
        font-size: 1.25rem !important;   
        letter-spacing: 0.15rem;         
        
        /* 微調清晰陰影，讓筆畫更銳利乾淨 */
        text-shadow: 1px 1px 3px rgba(0,0,0,0.8); 
    }
    
    /* --- 模組區塊樣式 --- */
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
st.markdown('<h1 class="main-title">⚡華南高商 試務組 AI 智能輔助平台⚡</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Intelligent Examination Administration Ecosystem</p>', unsafe_allow_html=True)

# ==========================================
# 3. 吉卜力風校園建築物 Banner (純 HTML 與定位技術)
# ==========================================
image_path = os.path.join("assets", "school_ghibli.png")
img_base64 = get_image_base64(image_path)

if img_base64:
    banner_html = f'''
        <div class="hero-banner-container">
            <img src="data:image/png;base64,{img_base64}" class="hero-banner">
            <div class="banner-overlay-text">攝於115年夏</div>
        </div>
    '''
    st.markdown(banner_html, unsafe_allow_html=True)
else:
    st.warning(f"⚠️ 找不到吉卜力風 Banner 圖片。請確認檔案放置於 `assets/school_ghibli.png`。")

# ==========================================
# 4. 核心工具矩陣入口 (2x2 模組排列)
# ==========================================
st.markdown("### ⚙️ 核心自動化模組 /")
st.write("")

# --- 第一排：監考與補考 ---
row1_col1, row1_col2 = st.columns(2, gap="large")

with row1_col1:
    with st.container(border=True):
        st.markdown("### 📅 段考試務智能輔助系統")
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
        st.page_link("pages/1_段考監考智能輔助系統.py", label="啟動段考排班作業 →", icon="🚀")

with row1_col2:
    with st.container(border=True):
        st.markdown("### 📝 補考作業智能輔助系統")
        st.markdown("""
        <p class="module-desc">
        一鍵清洗多科補考學生複雜數據，智慧重組考場與座位配置。
        </p>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        - 🧼 **智慧數據清洗**：自動剃除幽靈名單與全半形錯字干擾。
        - 🏫 **多重考場分流**：自動計算最大考場容量，防範座位重疊衝突。
        - 🖨️ **試務報表輸出**：自動生成個人補考通知單與考場對照表。
        """)
        st.write("")
        st.page_link("pages/2_補考作業智能輔助系統.py", label="啟動補考處理作業 →", icon="⚡")

st.write("") # 增加排與排之間的間距

# --- 第二排：教師甄選與模擬考 ---
row2_col1, row2_col2 = st.columns(2, gap="large")

with row2_col1:
    with st.container(border=True):
        st.markdown("### 👨‍🏫 教師甄選智能輔助系統")
        st.markdown("""
        <p class="module-desc">
        自動化排定口試與面試時程，場地、排版、蓋章一鍵直出零失誤。
        </p>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        - 🎙️ **時程自動排定**：無縫接軌口試與試教，智慧避開衝堂與重疊。
        - 📋 **場地動態對位**：自動分配各科休息室與試場，資料無縫串接。
        - 🖨️ **一鍵直出報表**：自動生成紅字警語且蓋妥官印之Word公告單。
        """)
        st.write("")
        # 🌟 質感修復：標題為 👨‍🏫，按鈕使用最具行動導向的 🎯 圖示
        st.page_link("pages/3_教師甄選智能輔助系統.py", label="啟動教師甄選作業 →", icon="🎯")

with row2_col2:
    with st.container(border=True):
        st.markdown("### 📊 模考調查智能輔助系統")
        st.markdown("""
        <p class="module-desc">
        數位化統整各科別模擬考報名意願，自動精算測驗費用與收據報表。
        </p>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        - 📑 **科別智慧調查**：快速整合各班級與類科的模擬考參與意願。
        - 💰 **金額自動統計**：精準計算各班、各科的收費總計與跨科差額。
        - 📈 **圖表化清單**：自動匯出清晰的總務處繳費單與書商訂卷清單。
        """)
        st.write("")
        # 🌟 質感修復：標題為 📊，按鈕使用代表數據全面啟動的 📈 圖示
        st.page_link("pages/4_模考調查智能輔助系統.py", label="啟動模擬考調查作業 →", icon="📈")

# ==========================================
# 5. 底部沉穩頁尾
# ==========================================
st.divider()
st.markdown(
    """
    <div style="text-align: center; color: #bbb; font-size: 0.8rem; letter-spacing: 0.05rem;">
        © 2026 國立華南高商 教務處試務組 劉耀中老師監製· Data-Driven Examination Administration
    </div>
    """, 
    unsafe_allow_html=True
)
