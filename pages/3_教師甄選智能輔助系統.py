import streamlit as st
import pandas as pd
import datetime
import io
import os
import traceback

try:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, Cm, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==========================================
# 1. 網頁頁面配置與記憶體初始化
# ==========================================
st.set_page_config(page_title="教甄智能排程系統", page_icon="🏫", layout="wide")
st.title("🏫 教務處-教師甄選智能排程系統 (簽到表與評分表優化旗艦版)")
st.info("💡 終極優化：已修正簽到表右下角承辦人與印章靠右對齊，並大幅調寬、加高實作評分表之委員評語欄位，方便現場手寫紀錄！")

if not HAS_DOCX:
    st.error("🚨 偵測到系統未安裝 `python-docx` 套件！無法產出直出版 Word。請在 requirements.txt 中加入 `python-docx`。")

# 初始化系統記憶體
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'excel_data' not in st.session_state:
    st.session_state.excel_data = None
if 'word_data' not in st.session_state:
    st.session_state.word_data = None
if 'df_preview' not in st.session_state:
    st.session_state.df_preview = None

# ==========================================
# 0. 側邊欄：試務資源與印章設定
# ==========================================
st.sidebar.title("📥 試務資源下載")
template_filename = "114第1次代理教師甄選各科預定流程時間表[最新版]1140606.doc"
file_path = template_filename
if not os.path.exists(file_path):
    if os.path.exists(f"../{template_filename}"):
        file_path = f"../{template_filename}"
    elif os.path.exists(f"pages/{template_filename}"):
        file_path = f"pages/{template_filename}"

if os.path.exists(file_path):
    with open(file_path, "rb") as file:
        st.sidebar.download_button(
            label="📥 下載 Word 官方合併列印範本",
            data=file,
            file_name=template_filename,
            mime="application/msword",
            type="primary"
        )
else:
    st.sidebar.warning(f"⚠️ 找不到 `{template_filename}`，請確認已上傳至專案目錄。")

st.sidebar.divider()

st.sidebar.title("🔴 自動蓋章設定 (Word專用)")
st.sidebar.markdown("上傳您的「試務組印章.png」，系統將自動印在簽到表右下角。")
file_stamp = st.sidebar.file_uploader("上傳印章圖檔 (.png, .jpg)", type=['png', 'jpg', 'jpeg'])

# ==========================================
# 2. 華南教甄官方硬性時間矩陣資料庫
# ==========================================
TEACH_15_MATRIX = {
    1: ("09:40-09:55", "09:55-10:10"), 2: ("09:55-10:10", "10:10-10:25"),
    3: ("10:10-10:25", "10:25-10:40"), 4: ("10:25-10:40", "10:40-10:55"),
    5: ("10:40-10:55", "10:55-11:10"), 6: ("10:55-11:10", "11:10-11:25"),
    7: ("11:10-11:25", "11:25-11:40"), 8: ("11:25-11:40", "11:40-11:55"),
    9: ("13:10-13:25", "13:25-13:40"), 10: ("13:25-13:40", "13:40-13:55"),
    11: ("13:40-13:55", "13:55-14:10"), 12: ("13:55-14:10", "14:10-14:25"),
    13: ("14:10-14:25", "14:25-14:40"), 14: ("14:25-14:40", "14:40-14:55"),
    15: ("14:40-14:55", "14:55-15:10"), 16: ("14:55-15:10", "15:10-15:25"),
    17: ("15:10-15:25", "15:25-15:40"), 18: ("15:25-15:40", "15:40-15:55"),
    19: ("15:40-15:55", "15:55-16:10"), 20: ("15:55-16:10", "16:10-16:25"),
    21: ("16:10-16:25", "16:25-16:40"), 22: ("16:25-16:40", "16:40-16:55"),
    23: ("16:40-16:55", "16:55-17:10"), 24: ("16:55-17:10", "17:10-17:25"), 25: ("17:10-17:25", "17:25-17:40"),
}

ORAL_15_MATRIX = {
    1: {1: "10:10-10:20"},
    2: {1: "10:10-10:20", 2: "09:40-09:50"},
    3: {1: "10:10-10:20", 2: "09:40-09:50", 3: "09:55-10:05"},
    4: {1: "10:25-10:35", 2: "09:40-09:50", 3: "09:55-10:05", 4: "10:10-10:20"},
    5: {1: "10:40-10:50", 2: "09:40-09:50", 3: "09:55-10:05", 4: "10:10-10:20", 5: "10:25-10:35"},
    6: {1: "10:10-10:20", 2: "09:40-09:50", 3: "09:55-10:05", 4: "10:55-11:05", 5: "10:25-10:35", 6: "10:40-10:50"},
    7: {1: "10:10-10:20", 2: "09:40-09:50", 3: "09:55-10:05", 4: "11:10-11:20", 5: "10:25-10:35", 6: "10:40-10:50", 7: "10:55-11:05"},
    8: {1: "10:10-10:20", 2: "09:40-09:50", 3: "09:55-10:05", 4: "11:25-11:35", 5: "10:25-10:35", 6: "10:40-10:50", 7: "10:55-11:05", 8: "11:10-11:20"},
    9: {1: "10:10-10:20", 2: "09:40-09:50", 3: "09:55-10:05", 4: "10:55-11:05", 5: "10:25-10:35", 6: "10:40-10:50", 7: "11:40-11:50", 8: "11:10-11:20", 9: "11:25-11:35"},
    10: {
        1: "10:10-10:20", 2: "09:40-09:50", 3: "09:55-10:05", 4: "10:55-11:05", 5: "10:25-10:35", 6: "10:40-10:50", 
        7: "11:40-11:50", 8: "11:10-11:20", 9: "11:25-11:35", 10: "13:10-13:20", 11: "13:25-13:35", 12: "13:40-13:50", 
        13: "13:55-14:05", 14: "14:10-14:20", 15: "14:25-14:35", 16: "14:40-14:50", 17: "14:55-15:05", 18: "15:10-15:20",
        19: "15:25-15:35", 20: "15:40-15:50", 21: "15:55-16:05", 22: "16:10-16:20", 23: "16:25-16:35", 24: "16:40-16:50", 25: "16:55-17:05"
    }
}

TEACH_30_MATRIX = {
    1: ("09:40-09:55", "09:55-10:25"), 2: ("10:10-10:25", "10:25-10:55"),
    3: ("10:40-10:55", "10:55-11:25"), 4: ("11:10-11:25", "11:25-11:55"),
    5: ("13:10-13:25", "13:25-13:55"), 6: ("13:40-13:55", "13:55-14:25"),
    7: ("14:10-14:25", "14:25-14:55"), 8: ("14:40-14:55", "14:55-15:25"),
    9: ("15:10-15:25", "15:25-15:55"), 10: ("15:40-15:55", "15:55-16:25"),
    11: ("16:10-16:25", "16:25-16:55"), 12: ("16:40-16:55", "16:55-17:25"),
    13: ("17:10-17:25", "17:25-17:55"), 14: ("17:40-17:55", "17:55-18:25"), 15: ("18:10-18:25", "18:25-18:55"),
}

ORAL_30_MATRIX = {
    1: "13:40-13:50", 2: "13:55-14:05", 3: "14:10-14:20", 4: "14:21-14:31",
    5: "14:40-14:50", 6: "14:55-15:05", 7: "15:10-15:20", 8: "15:25-15:35",
    9: "15:40-15:50", 10: "15:55-16:05", 11: "16:10-16:20", 12: "16:25-16:35",
    13: "16:40-16:50", 14: "16:55-17:05", 15: "17:10-17:20",
}

# ==========================================
# 3. 介面佈局與參數設定
# ==========================================
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("📂 1. 上傳考生名單與場地")
    file_candidates = st.file_uploader("1️⃣ 上傳報考名單 (准考證號, 報考科目).xlsx", type=['xlsx'])
    file_venues = st.file_uploader("2️⃣ 上傳場地配置 (如 3.場地.xlsx)", type=['xlsx'])
    
    group_settings = []
    practical_subjects = []
    all_subjs = []
    venue_dict = {}
    
    if file_candidates:
        try:
            df_temp = pd.read_excel(file_candidates).fillna("")
            if '報考科目' in df_temp.columns:
                all_subjs = df_temp['報考科目'].unique().tolist()
                
                st.write("---")
                st.markdown("### 🤝 2. 設定合併口試群組")
                num_groups = st.number_input("欲建立的「合併口試組」數量：", min_value=0, max_value=5, value=0, step=1)
                already_assigned = set()
                for g_i in range(int(num_groups)):
                    available_options = [s for s in all_subjs if s not in already_assigned]
                    selected_for_g = st.multiselect(
                        f"選擇【合併口試群組 {g_i + 1}】的成員科目：",
                        options=available_options,
                        key=f"group_select_{g_i}"
                    )
                    if selected_for_g:
                        group_settings.append(selected_for_g)
                        already_assigned.update(selected_for_g)
                
                st.write("---")
                st.markdown("### 🛠️ 3. 設定實作學科 (切換 30 分鐘時間軸)")
                practical_subjects = st.multiselect("選擇包含實作的科目：", options=all_subjs)
                
        except Exception as e:
            st.error(f"讀取失敗: {e}")

with col2:
    st.subheader("📝 4. 甄選標題設定")
    c_year, c_session = st.columns(2)
    with c_year:
        academic_year = st.number_input("學年度 (如: 114)", min_value=100, max_value=200, value=114, step=1)
    with c_session:
        session_num = st.number_input("第幾次甄選 (如: 1)", min_value=1, max_value=50, value=1, step=1)
    
    st.write("---")
    st.subheader("📋 5. 系統輸出說明")
    st.markdown("""
    本系統現已成為**全自動試務產出中心**：
    
    1. **時間表生成**：表格標題與內容置中，欄寬加寬防編號換行。
    2. **考生簽到表**：自動生成各科獨立簽到簿，**承辦人簽章與試務組印章全自動向右靠齊**。
    3. **委員評分表**：區分實作與一般評分表，**評語紀錄格大幅度拉寬加高**，方便評審執筆。
    """)

st.divider()

# ==========================================
# 4. 核心排程與生成演算法
# ==========================================
if st.button("🚀 啟動排程與場地整合", type="primary", use_container_width=True):
    if not file_candidates or not file_venues:
        st.error("🚨 請確認【1️⃣報考名單】與【2️⃣場地設定】皆已上傳！")
    else:
        try:
            dynamic_title = f"{academic_year}學年度第{session_num}次代理教師甄選人員試教及口試時間表"
            
            df_venues = pd.read_excel(file_venues).fillna("")
            for _, row in df_venues.iterrows():
                subj = str(row.get('科目', '')).strip()
                if subj:
                    venue_dict[subj] = {
                        '休息室': str(row.get('休息室', '')).strip(),
                        '準備室': str(row.get('準備室', '')).strip(),
                        '試教': str(row.get('試教', '')).strip(),
                        '口試': str(row.get('口試', '')).strip()
                    }

            df_candidates = pd.read_excel(file_candidates).fillna("")
            df_candidates['准考證號'] = df_candidates['准考證號'].astype(str).str.strip()
            df_candidates = df_candidates[df_candidates['准考證號'] != ""]

            if df_candidates.empty:
                st.warning("⚠️ 攔截警告：名單內已無有效考生！")
                st.stop()

            final_processing_groups = []
            assigned_set = set()
            for g_subs in group_settings:
                if g_subs:
                    final_processing_groups.append({'type': 'merged', 'subjects': g_subs})
                    assigned_set.update(g_subs)
            for sub in all_subjs:
                if sub not in assigned_set:
                    final_processing_groups.append({'type': 'independent', 'subjects': [sub]})

            all_schedules = []
            
            for group in final_processing_groups:
                group_total_candidates = sum(len(df_candidates[df_candidates['報考科目'] == sub]) for sub in group['subjects'])
                global_idx = 1
                
                for s_idx, subject in enumerate(group['subjects']):
                    is_practical = subject in practical_subjects
                    df_sub = df_candidates[df_candidates['報考科目'] == subject]
                    candidates = df_sub.to_dict('records')
                    n_candidates = len(candidates)
                    if n_candidates == 0: continue
                    
                    for i in range(n_candidates):
                        cand = candidates[i]
                        sort_num = i + 1 
                        
                        if is_practical:
                            times_tuple = TEACH_30_MATRIX.get(sort_num, ("請手動調整", "請手動調整"))
                            oral_range = ORAL_30_MATRIX.get(global_idx, "請手動調整")
                        else:
                            times_tuple = TEACH_15_MATRIX.get(sort_num, ("請手動調整", "請手動調整"))
                            lookup_n = group_total_candidates if group_total_candidates <= 9 else 10
                            oral_range = ORAL_15_MATRIX.get(lookup_n, {}).get(global_idx, "請手動調整")
                            
                        all_schedules.append({
                            '報考科目': subject,
                            '准考證號': cand['准考證號'],
                            '排序': sort_num,
                            '準備時間': times_tuple[0],
                            '試教(實作)時間': times_tuple[1],
                            '口試時間': oral_range
                        })
                        global_idx += 1

            df_master = pd.DataFrame(all_schedules)
            
            # --- Excel 產出 ---
            df_merge = df_master.copy()
            df_merge = df_merge.rename(columns={'報考科目': '科目', '准考證號': '准考證', '試教(實作)時間': '試教時間'})
            
            df_merge.insert(1, '休息室', df_merge['科目'].apply(lambda x: venue_dict.get(x, {}).get('休息室', '未設定')))
            df_merge.insert(2, '準備室', df_merge['科目'].apply(lambda x: venue_dict.get(x, {}).get('準備室', '未設定')))
            df_merge.insert(3, '試教', df_merge['科目'].apply(lambda x: venue_dict.get(x, {}).get('試教', '未設定')))
            df_merge.insert(4, '口試', df_merge['科目'].apply(lambda x: venue_dict.get(x, {}).get('口試', '未設定')))
            df_merge = df_merge[['科目', '休息室', '準備室', '試教', '口試', '准考證', '排序', '準備時間', '試教時間', '口試時間']]
            
            merge_with_blanks = []
            prev_subj = None
            for idx, row in df_merge.iterrows():
                curr_subj = row['科目']
                if prev_subj is not None and curr_subj != prev_subj:
                    merge_with_blanks.append({col: "" for col in df_merge.columns})
                merge_with_blanks.append(row.to_dict())
                prev_subj = curr_subj
            df_merge_final = pd.DataFrame(merge_with_blanks)
            
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                df_master.to_excel(writer, index=False, sheet_name='試務中心總表')
                df_merge_final.to_excel(writer, index=False, sheet_name='合併列印專用')
                df_teach = df_master[['報考科目', '排序', '試教(實作)時間', '准考證號']].copy()
                df_teach.to_excel(writer, index=False, sheet_name='門口_試教實作表')
                df_int = df_master[['報考科目', '准考證號', '口試時間']].copy()
                df_int['開始時間'] = df_int['口試時間'].str[:5]
                df_int = df_int.sort_values(['報考科目', '開始時間']).drop(columns=['開始時間'])
                df_int.to_excel(writer, index=False, sheet_name='門口_口試表')
            
            st.session_state.excel_data = output_excel.getvalue()

            # --- Word 直出引擎 (三大表大滿貫整合) ---
            if HAS_DOCX:
                doc = docx.Document()
                
                # 【全域文字基礎設定】
                style = doc.styles['Normal']
                style.font.name = '標楷體'
                style._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '標楷體')
                style.font.size = Pt(16)
                
                # 遍歷科目依序產出
                for current_idx, subject in enumerate(all_subjs):
                    df_sub_sched = df_master[df_master['報考科目'] == subject]
                    if df_sub_sched.empty: continue
                    is_practical = subject in practical_subjects
                    v = venue_dict.get(subject, {})
                    
                    # ----------------------------------------------------
                    # 表格 A：【各科預定流程時間表】(配置專屬頁尾警語)
                    # ----------------------------------------------------
                    sec_sched = doc.sections[0] if current_idx == 0 else doc.add_section()
                    sec_sched.bottom_margin = Cm(3.0)
                    sec_sched.footer_distance = Cm(1.0)
                    
                    # 設置時間表專用紅色警語頁尾
                    footer_sched = sec_sched.footer
                    footer_sched.is_linked_to_previous = False
                    for p in list(footer_sched.paragraphs):
                        footer_sched._element.remove(p._element)
                    p_ft = footer_sched.add_paragraph()
                    p_ft.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    r_ft = p_ft.add_run("※試教及口試時間將依現場實際報到人數及試場情形作調整，請考生於各科指定之休息室等候叫號")
                    r_ft.font.name = '標楷體'
                    r_ft._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '標楷體')
                    r_ft.font.size = Pt(11)
                    r_ft.font.color.rgb = RGBColor(255, 0, 0)
                    
                    # 內文抬頭
                    p_title = doc.add_paragraph()
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_title = p_title.add_run(dynamic_title)
                    run_title.font.size = Pt(18)
                    run_title.bold = True
                    
                    p_sub = doc.add_paragraph()
                    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sub = p_sub.add_run(f"【{subject}】")
                    run_sub.font.size = Pt(16)
                    run_sub.bold = True
                    
                    doc.add_paragraph(f"考生休息室：{v.get('休息室', '未設定')}")
                    doc.add_paragraph(f"試教準備室：{v.get('準備室', '未設定')}")
                    doc.add_paragraph(f"試教場地：{v.get('試教', '未設定')}")
                    doc.add_paragraph(f"口試場地：{v.get('口試', '未設定')}")
                    
                    table = doc.add_table(rows=1, cols=5)
                    table.style = 'Table Grid'
                    table.autofit = False 
                    col_widths = [Cm(3.5), Cm(2.0), Cm(3.5), Cm(3.5), Cm(3.5)]
                    
                    hdr_cells = table.rows[0].cells
                    hdr_headers = ['甄選證號', '編號', '試教準備室', '試教', '口試']
                    for col_idx in range(5):
                        hdr_cells[col_idx].text = hdr_headers[col_idx]
                        hdr_cells[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hdr_cells[col_idx].width = col_widths[col_idx]
                        table.columns[col_idx].width = col_widths[col_idx]
                    
                    for _, cand in df_sub_sched.iterrows():
                        row_cells = table.add_row().cells
                        row_data = [str(cand['准考證號']), str(cand['排序']), str(cand['準備時間']), str(cand['試教(實作)時間']), str(cand['口試時間'])]
                        for col_idx in range(5):
                            row_cells[col_idx].text = row_data[col_idx]
                            row_cells[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            row_cells[col_idx].width = col_widths[col_idx]
                    
                    # ----------------------------------------------------
                    # 表格 B：【考生報到簽到表】(斷開警語頁尾、右下角簽章靠右)
                    # ----------------------------------------------------
                    sec_signin = doc.add_section()
                    sec_signin.footer.is_linked_to_previous = False
                    for p in list(sec_signin.footer.paragraphs):
                        sec_signin.footer._element.remove(p._element)
                        
                    p_title2 = doc.add_paragraph()
                    p_title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_title2 = p_title2.add_run(f"{academic_year}學年度第{session_num}次代理教師甄選人員簽到表")
                    run_title2.font.size = Pt(18)
                    run_title2.bold = True
                    
                    p_sub2 = doc.add_paragraph()
                    p_sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sub2 = p_sub2.add_run(f"【{subject}】")
                    run_sub2.font.size = Pt(16)
                    run_sub2.bold = True
                    
                    doc.add_paragraph(f"甄選試場場地：{v.get('試教', '未設定')} / {v.get('口試', '未設定')}")
                    doc.add_paragraph() 
                    
                    table_si = doc.add_table(rows=1, cols=3)
                    table_si.style = 'Table Grid'
                    table_si.autofit = False
                    col_widths_si = [Cm(5.0), Cm(2.5), Cm(8.5)]
                    
                    hdr_cells_si = table_si.rows[0].cells
                    hdr_headers_si = ['甄選證號', '編號', '考生親筆簽名欄']
                    for col_idx in range(3):
                        hdr_cells_si[col_idx].text = hdr_headers_si[col_idx]
                        hdr_cells_si[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hdr_cells_si[col_idx].width = col_widths_si[col_idx]
                        table_si.columns[col_idx].width = col_widths_si[col_idx]
                        
                    for _, cand in df_sub_sched.iterrows():
                        row_cells = table_si.add_row().cells
                        row_cells[0].text = str(cand['准考證號'])
                        row_cells[1].text = str(cand['排序'])
                        row_cells[2].text = "" 
                        table_si.rows[-1].height = Cm(1.3) # 調配舒適簽名高度
                        for col_idx in range(3):
                            row_cells[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            row_cells[col_idx].width = col_widths_si[col_idx]
                            
                    doc.add_paragraph()
                    doc.add_paragraph()
                    
                    # 【核心修正】：簽到表中右下角的承辦人簽名欄與印章全面「靠右側」
                    p_sign = doc.add_paragraph()
                    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    run_sign = p_sign.add_run("承辦人簽章：__________________ 　　\n\n")
                    run_sign.font.name = '標楷體'
                    run_sign.font.size = Pt(14)
                    
                    if file_stamp:
                        run_stamp = p_sign.add_run()
                        stamp_bytes = io.BytesIO(file_stamp.getvalue())
                        run_stamp.add_picture(stamp_bytes, width=Cm(3.5))
                        
                    # ----------------------------------------------------
                    # 表格 C：【評審委員評分表】(欄位極致放大、加高)
                    # ----------------------------------------------------
                    sec_eval = doc.add_section()
                    sec_eval.footer.is_linked_to_previous = False
                    for p in list(sec_eval.footer.paragraphs):
                        sec_eval.footer._element.remove(p._element)
                        
                    p_title3 = doc.add_paragraph()
                    p_title3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # 動態識別是否為實作科目
                    title_text = f"{academic_year}學年度第{session_num}次代理教師甄選【實作評分表】" if is_practical else f"{academic_year}學年度第{session_num}次代理教師甄選【試教及口試評分表】"
                    run_title3 = p_title3.add_run(title_text)
                    run_title3.font.size = Pt(18)
                    run_title3.bold = True
                    
                    p_sub3 = doc.add_paragraph()
                    p_sub3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sub3 = p_sub3.add_run(f"【{subject}】")
                    run_sub3.font.size = Pt(16)
                    run_sub3.bold = True
                    
                    doc.add_paragraph(f"評分專用場地：{v.get('試教', '未設定') if not is_practical else v.get('口試', '未設定')}")
                    doc.add_paragraph() 
                    
                    table_ev = doc.add_table(rows=1, cols=4)
                    table_ev.style = 'Table Grid'
                    table_ev.autofit = False
                    
                    # 【核心修正】：將「評分內容與委員評語」欄位寬度極致放大至 9.0cm
                    col_widths_ev = [Cm(3.5), Cm(1.5), Cm(9.0), Cm(2.0)] 
                    
                    hdr_cells_ev = table_ev.rows[0].cells
                    hdr_headers_ev = ['甄選證號', '編號', '評分內容與委員具體評語 (欄位已擴大)', '實得分數']
                    for col_idx in range(4):
                        hdr_cells_ev[col_idx].text = hdr_headers_ev[col_idx]
                        hdr_cells_ev[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hdr_cells_ev[col_idx].width = col_widths_ev[col_idx]
                        table_ev.columns[col_idx].width = col_widths_ev[col_idx]
                        
                    for _, cand in df_sub_sched.iterrows():
                        row_cells = table_ev.add_row().cells
                        row_cells[0].text = str(cand['准考證號'])
                        row_cells[1].text = str(cand['排序'])
                        row_cells[2].text = "" 
                        row_cells[3].text = "" 
                        
                        # 【核心修正】：強制拉高欄位高度至 2.5公分，完美留白方便現場紀錄
                        table_ev.rows[-1].height = Cm(2.5) 
                        
                        for col_idx in range(4):
                            row_cells[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            row_cells[col_idx].width = col_widths_ev[col_idx]
                            
                    doc.add_paragraph()
                    doc.add_paragraph()
                    
                    p_ev_sign = doc.add_paragraph()
                    p_ev_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    run_ev_sign = p_ev_sign.add_run("評審委員親筆簽章：______________________ 　　")
                    run_ev_sign.font.name = '標楷體'
                    run_ev_sign.font.size = Pt(14)

                output_word = io.BytesIO()
                doc.save(output_word)
                st.session_state.word_data = output_word.getvalue()
                
            st.session_state.df_preview = df_merge_final.head(15)
            st.session_state.excel_filename = f"{academic_year}學年度第{session_num}次_排程與場地整合表.xlsx"
            st.session_state.word_filename = f"{academic_year}學年度第{session_num}次_各科甄選試務大滿貫報表.docx"
            st.session_state.processed = True

        except Exception as e:
            st.error(f"發生錯誤: {e}")
            st.code(traceback.format_exc())

# ==========================================
# 5. 結果顯示區
# ==========================================
if st.session_state.processed:
    st.balloons()
    st.success("🎉 修正排版完美達成！簽到表印章與簽名已靠右，評分表手寫格大幅拓寬加高。")
    
    c_d1, c_d2 = st.columns(2)
    with c_d1:
        st.download_button(
            label="📥 1. 下載 Excel 總表 (含合併列印專用)",
            data=st.session_state.excel_data,
            file_name=st.session_state.excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
    with c_d2:
        if HAS_DOCX and st.session_state.word_data:
            st.download_button(
                label="📥 2. 下載 Word 各科試務大滿貫報表 (含時間表/簽到簿/放大評分表)",
                data=st.session_state.word_data,
                file_name=st.session_state.word_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )

    st.write("👀 **專屬 Word 對接資料預覽：**")
    st.dataframe(st.session_state.df_preview)
