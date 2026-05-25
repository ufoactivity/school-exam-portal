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
st.title("🏫 教務處-教師甄選智能排程系統 (排版置中旗艦版)")
st.info("💡 終極優化：頁尾警語已實裝「分段多色雙字體」強調設計，且已恢復「自動偵測預設印章」功能，免重複上傳！")

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

# 【恢復功能】：自動偵測預設印章邏輯
default_stamp_name = "試務組印章.png"
stamp_path = default_stamp_name
has_default_stamp = False
file_stamp = None

if os.path.exists(stamp_path):
    has_default_stamp = True
elif os.path.exists(f"../{default_stamp_name}"):
    stamp_path = f"../{default_stamp_name}"
    has_default_stamp = True
elif os.path.exists(f"pages/{default_stamp_name}"):
    stamp_path = f"pages/{default_stamp_name}"
    has_default_stamp = True

if has_default_stamp:
    st.sidebar.success(f"✅ 系統已自動載入預設印章：\n`{default_stamp_name}`")
    with st.sidebar.expander("🔄 想要臨時更換印章？點此手動上傳"):
        file_stamp = st.file_uploader("上傳新印章圖檔 (.png, .jpg)", type=['png', 'jpg', 'jpeg'])
else:
    st.sidebar.markdown("上傳您的「試務組印章.png」，系統將自動印在每頁右下角。")
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
    
    1. **排版優化**：表格標題與內容皆已「全面置中」，「編號」欄寬微調加寬防換行。
    2. **完美頁尾設計**：頁尾紅字與印章距離底端 1cm，且警語已實裝重點放大標紅功能。
    3. **雙軌下載**：提供手動 Excel 套印與 Word 一鍵直出雙功能。
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

            # --- Word 直出引擎 (置中排版 + 自動蓋章) ---
            if HAS_DOCX:
                doc = docx.Document()
                
                # 【全域設定】
                section = doc.sections[0]
                section.footer_distance = Cm(1.0)
                section.bottom_margin = Cm(3.0)
                
                style = doc.styles['Normal']
                style.font.name = '標楷體'
                style._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '標楷體')
                style.font.size = Pt(16)
                
                # 【頁尾引擎】
                footer = section.footer
                for p in footer.paragraphs:
                    p._element.getparent().remove(p._element)
                
                footer_table = footer.add_table(rows=1, cols=2, width=Cm(16.0))
                footer_table.autofit = False
                
                cell_left = footer_table.rows[0].cells[0]
                cell_left.width = Cm(11.5)
                cell_right = footer_table.rows[0].cells[1]
                cell_right.width = Cm(4.5)
                
                p_footer_text = cell_left.paragraphs[0]
                p_footer_text.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                # 第一段：黑色，14pt
                run_1 = p_footer_text.add_run("※試教及口試時間")
                run_1.font.name = '標楷體'
                run_1._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '標楷體')
                run_1.font.size = Pt(14)
                run_1.font.color.rgb = RGBColor(0, 0, 0)
                
                # 第二段：紅色，16pt
                run_2 = p_footer_text.add_run("將依現場實際報到人數及試場情形作調整，")
                run_2.font.name = '標楷體'
                run_2._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '標楷體')
                run_2.font.size = Pt(16)
                run_2.font.color.rgb = RGBColor(255, 0, 0)
                
                # 第三段：黑色，14pt
                run_3 = p_footer_text.add_run("請考生於各科指定之休息室等候叫號")
                run_3.font.name = '標楷體'
                run_3._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '標楷體')
                run_3.font.size = Pt(14)
                run_3.font.color.rgb = RGBColor(0, 0, 0)
                
                # 【恢復功能】：決定印章來源並套印
                stamp_source = None
                if file_stamp:
                    stamp_source = io.BytesIO(file_stamp.getvalue())
                elif has_default_stamp:
                    stamp_source = stamp_path
                
                if stamp_source:
                    p_stamp = cell_right.paragraphs[0]
                    p_stamp.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
                    run_stamp = p_stamp.add_run()
                    run_stamp.add_picture(stamp_source, width=Cm(4.0)) 
                
                # 內容生成
                for subject in all_subjs:
                    df_sub_sched = df_master[df_master['報考科目'] == subject]
                    if df_sub_sched.empty: continue
                    
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
                    
                    v = venue_dict.get(subject, {})
                    doc.add_paragraph(f"考生休息室：{v.get('休息室', '未設定')}")
                    doc.add_paragraph(f"試教準備室：{v.get('準備室', '未設定')}")
                    doc.add_paragraph(f"試教場地：{v.get('試教', '未設定')}")
                    doc.add_paragraph(f"口試場地：{v.get('口試', '未設定')}")
                    
                    # 強制置中與欄寬微調
                    table = doc.add_table(rows=1, cols=5)
                    table.style = 'Table Grid'
                    table.autofit = False 
                    
                    # 將編號加寬至 2.0cm 確保 16pt 不會換行，其餘平均分配 (總計約 16cm)
                    col_widths = [Cm(3.5), Cm(2.0), Cm(3.5), Cm(3.5), Cm(3.5)]
                    
                    table.style.font.name = '標楷體'
                    table.style._element.rPr.rFonts.set(docx.oxml.ns.qn('w:eastAsia'), '標楷體')
                    table.style.font.size = Pt(16)
                    
                    # 設定表頭與置中
                    hdr_cells = table.rows[0].cells
                    hdr_headers = ['甄選證號', '編號', '試教準備室', '試教', '口試']
                    for col_idx in range(5):
                        hdr_cells[col_idx].text = hdr_headers[col_idx]
                        hdr_cells[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER 
                        hdr_cells[col_idx].width = col_widths[col_idx]
                        table.columns[col_idx].width = col_widths[col_idx]
                    
                    # 填入考生資料並全部置中
                    for _, cand in df_sub_sched.iterrows():
                        row_cells = table.add_row().cells
                        
                        row_data = [
                            str(cand['准考證號']),
                            str(cand['排序']),
                            str(cand['準備時間']),
                            str(cand['試教(實作)時間']),
                            str(cand['口試時間'])
                        ]
                        
                        for col_idx in range(5):
                            row_cells[col_idx].text = row_data[col_idx]
                            row_cells[col_idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            row_cells[col_idx].width = col_widths[col_idx]
                    
                    doc.add_page_break()
                    
                output_word = io.BytesIO()
                doc.save(output_word)
                st.session_state.word_data = output_word.getvalue()
                
            st.session_state.df_preview = df_merge_final.head(15)
            st.session_state.excel_filename = f"{academic_year}學年度第{session_num}次_排程與場地整合表.xlsx"
            st.session_state.word_filename = f"{academic_year}學年度第{session_num}次_各科蓋章公告表.docx"
            st.session_state.processed = True

        except Exception as e:
            st.error(f"發生錯誤: {e}")
            st.code(traceback.format_exc())

# ==========================================
# 5. 結果顯示區
# ==========================================
if st.session_state.processed:
    st.balloons()
    st.success("🎉 排版完美達成！頁尾警語已實裝分色字體強化，表格標題與內容皆已置中。")
    
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
                label="📥 2. 下載 Word 各科公告時間表 (置中排版蓋章版)",
                data=st.session_state.word_data,
                file_name=st.session_state.word_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )

    st.write("👀 **專屬 Word 對接資料預覽：**")
    st.dataframe(st.session_state.df_preview)
