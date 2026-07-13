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
# 📌 階段一所需套件 (python-docx) 與 排版模組
# ==========================================
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.shared import Pt
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==========================================
# 1. 網頁頁面配置與全局狀態記憶體初始化
# ==========================================
st.set_page_config(page_title="段考試務全能系統", page_icon="🏫", layout="wide")
st.title("🏫 試務組 - 段考試務全能系統 (旗艦整合版)")
st.info("💡 雲端防護升級：已實裝「Excel 幽靈空白列淨化」與「AI 單執行緒限制」，確保雙階段作業在雲端執行不當機！")

# --- 狀態記憶體初始化 (Session State) ---
if 'results_p2' not in st.session_state:
    st.session_state['results_p2'] = None
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# [更新] 階段一雙版本記憶體
if 'docx_data_p1_print' not in st.session_state:
    st.session_state['docx_data_p1_print'] = None
if 'docx_data_p1_msg' not in st.session_state:
    st.session_state['docx_data_p1_msg'] = None
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
# 2. 輔助功能定義
# ==========================================
def to_excel_bytes(df, header_df=None):
    output = io.BytesIO()
    if header_df is not None:
        final_out = pd.concat([header_df, df], ignore_index=True)
    else:
        final_out = df
    
    final_out = final_out.fillna("")
    
    for col in final_out.columns:
        final_out[col] = final_out[col].apply(lambda x: f" {x}" if isinstance(x, str) and (x.startswith("=") or x.startswith("-")) else x)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_out.to_excel(writer, index=False, header=False)
    return output.getvalue()

def normalize_cls(c):
    if pd.isna(c) or c is None: return ""
    s = str(c).strip().replace('ㄧ', '一').replace(' ', '').replace(' ', '')
    s = s.translate(str.maketrans('１２３４５６７８９０', '1234567890'))
    return s

def normalize_subject(s):
    s = str(s).strip().replace(' ', '').replace(' ', '')
    s = s.replace('國文', '國語文').replace('英文', '英語文')
    return s

def get_ai_date_str(j, day_starts, ai_date_strs):
    day_idx = 0
    for idx, ds in enumerate(day_starts):
        if j >= ds: day_idx = idx
    return ai_date_strs[min(day_idx, len(ai_date_strs)-1)]

def extract_period_num(s):
    if pd.isna(s): return -1
    s = str(s).strip()
    if any(k in s for k in ['月', '日', '年', '表', '華南', '期中', '次數', '日期']): return -1
    cn_to_num = {'一':'1', '二':'2', '三':'3', '四':'4', '五':'5', 
                 '六':'6', '七':'7', '八':'8', '九':'9', '十':'10', 
                 '１':'1', '２':'2', '３':'3', '４':'4', '５':'5', '６':'6', '７':'7'}
    for k, v in cn_to_num.items(): s = s.replace(k, v)
    nums = re.findall(r'\d+', s)
    if nums:
        p = int(nums[0])
        if 1 <= p <= 15: return p
    return -1

def matches_date(val_str, d_date):
    if not d_date: return False
    s = str(val_str).replace(' ', '').replace('nan', '')
    if not s: return False
    d1 = d_date.strftime('%Y-%m-%d')
    d2 = d_date.strftime('%m-%d')
    d3 = d_date.strftime('%Y/%m/%d')
    d4 = d_date.strftime('%m月%d日')
    d5 = d_date.strftime('%d日')
    d6 = str(d_date.day) + '日'
    return any(x in s for x in [d1, d2, d3, d4, d5, d6])

# ==========================================
# 3. 核心介面佈局：利用 Tabs 分割兩個階段
# ==========================================
tab1, tab2 = st.tabs(["📑 階段一：試卷催繳通知單", "📅 階段二：段考監考智能輔助系統"])

# ---------------------------------------------------------
# 【階段一：試卷催繳通知單】
# ---------------------------------------------------------
with tab1:
    st.subheader("📑 階段一：試卷催繳通知單自動生成系統")
    st.markdown("上傳包含催繳名單的 Excel，**選擇對應的工作表 (考試類型)**，系統會自動產出專屬的 Word 通知單。")
    
    if not HAS_DOCX:
        st.error("🚨 偵測到系統未安裝 `python-docx` 套件！請在環境中安裝 `python-docx`。")

    col1_p1, col2_p1 = st.columns([1, 1], gap="large")
    
    with col1_p1:
        st.markdown("##### ⚙️ 參數設定")
        deadline = st.text_input("📅 繳交截止日", value="6/26", help="例如：6/26", key="p1_deadline")
        sender_name = st.text_input("✍️ 發送人署名", value="試務組 耀中", key="p1_sender")
        
    with col2_p1:
        st.markdown("##### 📂 資料上傳與選擇")
        uploaded_file_p1 = st.file_uploader("請上傳「試卷催繳名單」(Excel 格式)", type=["xlsx", "xls"], key="p1_uploader")
        
        selected_sheet_p1 = None
        if uploaded_file_p1 is not None:
            try:
                excel_file_p1 = pd.ExcelFile(uploaded_file_p1)
                sheet_names = excel_file_p1.sheet_names
                selected_sheet_p1 = st.selectbox("👇 請選擇工作表 (系統將以此作為「考試類型」)：", sheet_names, key="p1_sheet_select")
            except Exception as e:
                st.error("無法讀取 Excel 檔案，請確認檔案是否毀損。")

    if st.button("🚀 一鍵產出雙版本催繳通知單", use_container_width=True, type="primary", key="btn_p1"):
        if not HAS_DOCX:
            st.error("系統缺少 python-docx 套件，無法執行。")
        elif uploaded_file_p1 is None or selected_sheet_p1 is None:
            st.warning("⚠️ 老師，請先上傳名單檔案，並選擇要處理的工作表喔！")
        else:
            try:
                # 【記憶體防護】：加上 dropna
                df = pd.read_excel(uploaded_file_p1, sheet_name=selected_sheet_p1).dropna(how='all')
                required_cols = ['年級', '科目名稱', '姓名']
                missing_cols = [col for col in required_cols if col not in df.columns]
                
                if missing_cols:
                    st.error(f"🚨 上傳的檔案缺少必備欄位：{', '.join(missing_cols)}。請檢查欄位名稱是否吻合。")
                else:
                    df['姓名'] = df['姓名'].astype(str).str.strip().replace('nan', '')
                    df['科目名稱'] = df['科目名稱'].astype(str).str.strip().replace('nan', '')
                    df['年級'] = df['年級'].astype(str).str.strip().replace('nan', '')
                    df = df[df['姓名'] != '']
                    
                    doc_print = Document()
                    doc_msg = Document()
                    
                    grouped = df.groupby('姓名')
                    
                    for idx, (name, group) in enumerate(grouped):
                        exam_type = selected_sheet_p1
                        count = len(group)
                        
                        table = doc_print.add_table(rows=1, cols=1)
                        table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        table.style = 'Table Grid'
                        
                        cell = table.cell(0, 0)
                        p_title_print = cell.paragraphs[0]
                        p_title_print.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        run_title_print = p_title_print.add_run(f"【{exam_type}】催繳試卷通知單")
                        run_title_print.bold = True
                        run_title_print.font.size = Pt(20) 
                        doc_print.add_paragraph()
                        
                        p_title_msg = doc_msg.add_paragraph()
                        p_title_msg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run_title_msg = p_title_msg.add_run(f"【{exam_type}】催繳試卷通知單")
                        run_title_msg.bold = True
                        run_title_msg.font.size = Pt(16)
                        doc_msg.add_paragraph()

                        for doc in [doc_print, doc_msg]:
                            doc.add_paragraph(f"{name} 老師您好:\n")
                            doc.add_paragraph(f"{exam_type}試卷繳交截止日 {deadline} 已過，溫馨提醒您尚有 {count} 份試卷未繳:\n")
                            
                            for grade, grade_group in group.groupby('年級'):
                                doc.add_paragraph(f"[{grade}年級]")
                                for i, (_, row) in enumerate(grade_group.iterrows(), 1):
                                    doc.add_paragraph(f"  {i}. 科目: {row['科目名稱']}")
                            
                            doc.add_paragraph(f"\n{sender_name}")
                        
                        if idx < len(grouped) - 1:
                            doc_print.add_page_break()
                            doc_msg.add_paragraph("\n" + "=" * 40 + "\n")
                    
                    out_stream_print = io.BytesIO()
                    doc_print.save(out_stream_print)
                    st.session_state['docx_data_p1_print'] = out_stream_print.getvalue()
                    
                    out_stream_msg = io.BytesIO()
                    doc_msg.save(out_stream_msg)
                    st.session_state['docx_data_p1_msg'] = out_stream_msg.getvalue()
                    
                    st.session_state['processed_p1'] = True
                    
            except Exception as e:
                st.error(f"發生未預期錯誤: {e}")
                st.code(traceback.format_exc())

    if st.session_state['processed_p1'] and st.session_state['docx_data_p1_print']:
        st.success(f"✅ 完美達成！已為您同步產出【紙本列印版】與【訊息複製版】。")
        c1, c2 = st.columns([1, 1])
        with c1:
            st.download_button(
                label=f"🖨️ 點我下載：紙本列印版 (自動換頁加框線)",
                data=st.session_state['docx_data_p1_print'],
                file_name=f"{selected_sheet_p1}試卷催繳通知單_紙本列印版.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary",
                key="dl_btn_p1_print"
            )
        with c2:
            st.download_button(
                label=f"💬 點我下載：訊息複製版 (連續文字便於私訊)",
                data=st.session_state['docx_data_p1_msg'],
                file_name=f"{selected_sheet_p1}試卷催繳通知單_訊息複製版.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="secondary",
                key="dl_btn_p1_msg"
            )


# ---------------------------------------------------------
# 【階段二：段考監考智能輔助系統】
# ---------------------------------------------------------
with tab2:
    st.subheader("📅 階段二：段考監考智能輔助系統 (終極完全體)")
    st.info("💡 終極升級：實裝「雙重智慧檢核機制」！AI 將在匯出總表全自動對帳，精準抓出人力缺口。")

    col1_p2, col2_p2 = st.columns([1, 1], gap="large")

    with col1_p2:
        st.markdown("##### 📂 1. 上傳排考與標籤資料")
        file_quota = st.file_uploader("1️⃣ 監考堂數.xlsx", type=['xlsx'], key=f"f1_{st.session_state['uploader_key']}")
        file_list = st.file_uploader("2️⃣ 監考名單.xlsx", type=['xlsx'], key=f"f2_{st.session_state['uploader_key']}")
        file_type = st.file_uploader("3️⃣ 監考類型總數.xlsx", type=['xlsx'], key=f"f3_{st.session_state['uploader_key']}")
        file_pub = st.file_uploader("4️⃣ 監考總表公布版.xlsx (範本)", type=['xlsx'], key=f"f4_{st.session_state['uploader_key']}")
        file_assign = st.file_uploader("5️⃣ 監考一覽表.xlsx (班級分配範本)", type=['xlsx'], key=f"f5_{st.session_state['uploader_key']}")
        st.write("---")
        file_course = st.file_uploader("6️⃣ 配課表.xlsx (多工作表)", type=['xlsx'], key=f"f6_{st.session_state['uploader_key']}")
        file_label = st.file_uploader("7️⃣ 標籤列印.xlsx (試卷袋範本)", type=['xlsx'], key=f"f7_{st.session_state['uploader_key']}")

    with col2_p2:
        st.markdown("##### ⚙️ 2. 考試設定與特許名單")
        selected_sheet = None
        if file_quota:
            xls = pd.ExcelFile(file_quota)
            selected_sheet = st.selectbox("👇 選擇考試項目：", xls.sheet_names, key="p2_sheet_select")
        
        flex_names = []
        teacher_list = []
        if file_list:
            # 【記憶體防護】：讀取名單時加上 dropna，過濾幽靈空白列
            temp_df = pd.read_excel(file_list, header=None).dropna(how='all').fillna("")
            for c in range(5):
                try:
                    lst = temp_df.iloc[2:, c].astype(str).str.strip().tolist()
                    lst = [t for t in lst if t != "" and t != "nan" and not str(t).isdigit()]
                    if len(lst) > 10:
                        teacher_list = lst; break
                except: pass
            flex_names = st.multiselect("🛡️ 優先時數不大於名單：", options=teacher_list, key="p2_flex_names")

        class_list = []
        if file_assign:
            # 【記憶體防護】：讀取一覽表時加上 dropna
            df_assign_temp = pd.read_excel(file_assign, header=None).dropna(how='all').fillna("")
            raw_list = df_assign_temp.iloc[:, 0].astype(str).str.strip().tolist()
            class_names_raw = [x for x in raw_list if x and not any(bad in x for bad in ["班級", "日期", "節次", "星期", "一覽表", "總表", "華南", "期中考", "註"])]
            class_list = [normalize_cls(c) for c in class_names_raw]

        st.write("")
        c_d0, c_d1, c_d2 = st.columns(3)
        with c_d0:
            has_manual = st.checkbox("📌 包含手排日", value=True, key="p2_has_manual")
            if has_manual: d0_date = st.date_input("📅 手排日期", datetime.now(), key="p2_d0_date")
            else: d0_date = None
                
        with c_d1: d1_date = st.date_input("📅 AI Day1", datetime.now(), key="p2_d1_date")
        with c_d2: d2_date = st.date_input("📅 AI Day2", datetime.now(), key="p2_d2_date")
        
        st.write("---")
        st.markdown("#### ⏳ 兼課教師可用時段精確鎖定")
        st.info("💡 設定兼課/代課老師**允許排考**的日期與節次（若不限制節次請留空）。AI 會自動避開其餘所有時段。")
        
        edited_time_df = st.data_editor(
            st.session_state.time_rules, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={
                "老師": st.column_config.SelectboxColumn("👨‍🏫 指定老師", options=teacher_list if teacher_list else [""], required=False),
                "允許日期": st.column_config.SelectboxColumn("📅 允許日期", options=["無限制", "僅 Day 1", "僅 Day 2"], default="無限制"),
                "允許節次": st.column_config.TextColumn("⏰ 允許節次 (如: 1,2,3)", help="請輸入阿拉伯數字並以逗號分隔，留空代表該日全天皆可。")
            },
            key="p2_time_editor"
        )

        st.write("---")
        st.markdown("#### 🎯 特定班級與老師綁定")

        file_bind = st.file_uploader("📥 [選填] 匯入既有綁定名單 (.xlsx)", type=['xlsx'], key=f"f_bind_{st.session_state['uploader_key']}")
        if file_bind and st.session_state.last_bind_file != file_bind.name:
            try:
                df_bind_up = pd.read_excel(file_bind).dropna(how='all')
                if "老師" in df_bind_up.columns and "班級" in df_bind_up.columns:
                    st.session_state.bind_rules = df_bind_up[["老師", "班級"]]
                    st.session_state.last_bind_file = file_bind.name
                    st.rerun() 
            except Exception as e:
                st.error("讀取失敗，請確認檔案是否為之前下載的格式。")

        edited_bind_df = st.data_editor(
            st.session_state.bind_rules, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={
                "老師": st.column_config.SelectboxColumn("👨‍🏫 指定老師", options=teacher_list if teacher_list else [""], required=False),
                "班級": st.column_config.SelectboxColumn("🏫 指定班級", options=class_list if class_list else [""], required=False)
            },
            key="p2_bind_editor"
        )
        
        bind_output = io.BytesIO()
        with pd.ExcelWriter(bind_output, engine='xlsxwriter') as writer:
            edited_bind_df.to_excel(writer, index=False, header=True)
        bind_excel_bytes = bind_output.getvalue()
        
        st.download_button(
            label="💾 儲存目前的綁定名單 (下載 .xlsx)",
            data=bind_excel_bytes,
            file_name="特定老師班級綁定名單.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="p2_dl_bind"
        )
        
        st.write("---")
        force_run = st.checkbox("⚠️ 忽略健檢警告，強制執行", key="p2_force_run")
        
        if st.button("🗑️ 清除所有設定 (僅限階段二)", use_container_width=True):
            st.session_state['results_p2'] = None
            st.session_state['uploader_key'] += 1
            if 'bind_rules' in st.session_state: del st.session_state['bind_rules']
            if 'time_rules' in st.session_state: del st.session_state['time_rules']
            if 'last_bind_file' in st.session_state: del st.session_state['last_bind_file']
            st.rerun()

    # ==========================================
    # 4. 核心演算法執行
    # ==========================================
    st.divider()

    if st.button("🚀 啟動終極全自動排班系統", type="primary", use_container_width=True, key="btn_p2"):
        if not all([file_quota, file_list, file_type, file_assign]):
            st.error("🚨 請至少確認【1, 2, 3, 5】號基礎檔案皆已上傳！")
        else:
            try:
                # --- 1. 讀取配額與名單 ---
                # 【記憶體防護】：加上 dropna
                df_quota = pd.read_excel(file_quota, sheet_name=selected_sheet).dropna(how='all').fillna("")
                quota_dict = {}
                for r in range(df_quota.shape[0]):
                    name = str(df_quota.iloc[r, 0]).strip()
                    try: q = int(float(str(df_quota.iloc[r, 1]).strip()))
                    except: q = 0
                    if name: quota_dict[name] = q
                
                # 【記憶體防護】：加上 dropna
                df_list_raw = pd.read_excel(file_list, header=None).dropna(how='all').fillna("")
                
                header_row_idx = 1
                for r in range(min(5, df_list_raw.shape[0])):
                    if any(k in str(df_list_raw.iloc[r, 1]).strip() for k in ["教師", "姓名", "老師"]):
                        header_row_idx = r; break
                
                teacher_col_idx = 1
                quota_col_in_list = 2
                
                period_cols = []
                for c in range(3, df_list_raw.shape[1]):
                    val = str(df_list_raw.iloc[header_row_idx, c]).strip()
                    if extract_period_num(val) != -1:
                        period_cols.append(c)
                
                if len(period_cols) < 1:
                    st.error("🚨 無法從檔案中辨識出節次數量，請確認第4欄之後有打上節次數字。")
                    st.stop()
                
                total_periods = len(period_cols)
                date_row_idx = header_row_idx - 1 if header_row_idx > 0 else 0
                
                date_row_s = df_list_raw.iloc[date_row_idx, :].replace(['nan', ''], np.nan).ffill().fillna("")
                
                manual_cols = []
                ai_period_cols = []
                
                for c in period_cols:
                    d_str = str(date_row_s[c])
                    if has_manual and d0_date and matches_date(d_str, d0_date):
                        manual_cols.append(c)
                    else:
                        ai_period_cols.append(c)
                    
                ai_periods = len(ai_period_cols)
                if ai_periods == 0:
                    st.error("🚨 扣除手排日後，找不到需要 AI 處理的欄位，請檢查日期設定！")
                    st.stop()
                    
                ai_period_nums = [extract_period_num(str(df_list_raw.iloc[header_row_idx, c])) for c in ai_period_cols]
                
                day_starts = [0]
                for j in range(1, ai_periods):
                    if ai_period_nums[j] <= ai_period_nums[j-1]: day_starts.append(j)

                # 【記憶體防護】：加上 dropna
                df_type = pd.read_excel(file_type, header=None).dropna(how='all').fillna("")
                req_matrix = {'△': [0]*ai_periods, '※': [0]*ai_periods}
                for i in range(len(df_type)):
                    row_name = str(df_type.iloc[i, 0]).strip()
                    if row_name in ['△', '※']:
                        req_list = []
                        for c in range(1, df_type.shape[1]):
                            v = str(df_type.iloc[i, c]).strip()
                            if v:
                                try: req_list.append(int(float(v)))
                                except: pass
                        
                        if len(req_list) >= total_periods:
                            req_matrix[row_name] = [req_list[period_cols.index(c)] for c in ai_period_cols]
                        else:
                            req_padded = (req_list + [0]*total_periods)[:total_periods]
                            req_matrix[row_name] = [req_padded[period_cols.index(c)] for c in ai_period_cols]

                ai_date_strs = [d1_date.strftime('%m月%d日'), d2_date.strftime('%m月%d日')]
                
                header_df = df_list_raw.iloc[0:header_row_idx+1].copy().astype(str).replace('nan', '')
                if has_manual and d0_date:
                    for mc in manual_cols:
                        header_df.iloc[date_row_idx, mc] = d0_date.strftime('%m月%d日')
                for j in range(ai_periods):
                    header_df.iloc[date_row_idx, ai_period_cols[j]] = get_ai_date_str(j, day_starts, ai_date_strs)
                
                df_list = df_list_raw.iloc[header_row_idx+1:].copy()
                # 【記憶體防護】：過濾名單中空字串與 NaN
                teachers = [str(x).strip() for x in df_list.iloc[:, teacher_col_idx] if pd.notna(x) and str(x).strip() != "" and str(x).strip() != "nan"]

                time_constraints = {}
                for _, row in edited_time_df.iterrows():
                    t_name = str(row['老師']).strip()
                    if t_name and t_name != 'None':
                        d_limit = str(row['允許日期']).strip()
                        p_limit_str = str(row['允許節次']).strip()
                        p_limits = [int(p) for p in re.findall(r'\d+', p_limit_str)] if p_limit_str else []
                        time_constraints[t_name] = {'day': d_limit, 'periods': p_limits}

                # --- 2. PuLP 運算 ---
                with st.spinner(f"🧠 實裝 8 大規則運算中 (已過濾幽靈資料，鎖定 {len(teachers)} 位有效教師)..."):
                    prob = pulp.LpProblem("Scheduling", pulp.LpMinimize)
                    vX = {}; vY = {}
                    for i in range(len(teachers)):
                        vX[i] = {}; vY[i] = {}
                        for j in range(ai_periods):
                            vX[i][j] = pulp.LpVariable(f"X_{i}_{j}", cat='Binary')
                            vY[i][j] = pulp.LpVariable(f"Y_{i}_{j}", cat='Binary')
                    
                    d1_idx = list(range(day_starts[0], day_starts[1])) if len(day_starts) > 1 else list(range(ai_periods))
                    d2_idx = list(range(day_starts[1], ai_periods)) if len(day_starts) > 1 else []

                    penalty = 0
                    for i, t in enumerate(teachers):
                        tgt = int(quota_dict.get(t, 0))
                        act = pulp.lpSum([vX[i][k] + vY[i][k]*2 for k in range(ai_periods)])
                        
                        dfct_pos = pulp.LpVariable(f"dfct_pos_{i}", 0)
                        dfct_neg = pulp.LpVariable(f"dfct_neg_{i}", 0)
                        prob += act + dfct_neg - dfct_pos == tgt
                        
                        penalty += (dfct_pos + dfct_neg) * 500
                        if t in flex_names: penalty -= dfct_neg * 400
                        
                        is_time_constrained = t in time_constraints
                        if is_time_constrained:
                            tc = time_constraints[t]
                            for j in range(ai_periods):
                                is_day1 = j in d1_idx
                                is_day2 = j in d2_idx
                                
                                if tc['day'] == '僅 Day 1' and is_day2:
                                    prob += vX[i][j] == 0; prob += vY[i][j] == 0
                                elif tc['day'] == '僅 Day 2' and is_day1:
                                    prob += vX[i][j] == 0; prob += vY[i][j] == 0
                                    
                                if tc['periods']:
                                    if ai_period_nums[j] not in tc['periods']:
                                        prob += vX[i][j] == 0; prob += vY[i][j] == 0
                        
                        if tgt >= 5 and len(day_starts) >= 2 and not is_time_constrained:
                            prob += pulp.lpSum([vX[i][j] + vY[i][j] for j in d1_idx]) >= 1
                            prob += pulp.lpSum([vX[i][j] + vY[i][j] for j in d2_idx]) >= 1
                            prob += pulp.lpSum([vX[i][j] for j in range(ai_periods)]) >= 1
                            prob += pulp.lpSum([vY[i][j] for j in range(ai_periods)]) >= 1

                        for j in range(ai_periods):
                            prob += vX[i][j] + vY[i][j] <= 1
                            cell_val = str(df_list.iloc[i, ai_period_cols[j]]).strip()
                            if cell_val != "" and cell_val != "nan":
                                prob += vX[i][j] == 0; prob += vY[i][j] == 0
                                
                            if ai_period_nums[j] in [3, 5]:
                                prob += vX[i][j] == 0
                                
                        for j in range(ai_periods - 1):
                            if ai_period_nums[j] == 1 and ai_period_nums[j+1] == 2:
                                prob += vX[i][j+1] >= vY[i][j]
                                
                    for j in range(ai_periods):
                        req_d = req_matrix['△'][j]
                        req_s = req_matrix['※'][j]
                        slk_d_pos = pulp.LpVariable(f"slkd_pos_{j}", 0)
                        slk_d_neg = pulp.LpVariable(f"slkd_neg_{j}", 0)
                        prob += pulp.lpSum([vX[i][j] for i in range(len(teachers))]) + slk_d_neg - slk_d_pos == req_d
                        penalty += (slk_d_pos + slk_d_neg) * 10000
                        
                        slk_s_pos = pulp.LpVariable(f"slks_pos_{j}", 0)
                        slk_s_neg = pulp.LpVariable(f"slks_neg_{j}", 0)
                        prob += pulp.lpSum([vY[i][j] for i in range(len(teachers))]) + slk_s_neg - slk_s_pos == req_s
                        penalty += (slk_s_pos + slk_s_neg) * 10000
                        
                    prob += penalty
                    
                    # 【防崩潰保險】：強制套用單執行緒限制與時限
                    solver = pulp.PULP_CBC_CMD(timeLimit=45, msg=False, threads=1)
                    prob.solve(solver)

                    schedule_dict = {}
                    df_out_master = df_list.copy()
                    actual_matrix = {'△': [0]*ai_periods, '※': [0]*ai_periods}
                    
                    for i, t in enumerate(teachers):
                        res = []
                        df_out_master.iloc[i, quota_col_in_list] = int(quota_dict.get(t, 0))
                        
                        for j in range(ai_periods):
                            val = str(df_list.iloc[i, ai_period_cols[j]]).strip()
                            if val == "" or val == "nan":
                                if vX[i][j].varValue == 1: 
                                    val = "△"
                                    actual_matrix['△'][j] += 1
                                elif vY[i][j].varValue == 1: 
                                    val = "※"
                                    actual_matrix['※'][j] += 1
                                else: val = "" 
                            else:
                                if val == "△": actual_matrix['△'][j] += 1
                                elif val == "※": actual_matrix['※'][j] += 1
                                    
                            res.append(val)
                            df_out_master.iloc[i, ai_period_cols[j]] = val
                        schedule_dict[t] = res

                # ======== 👇 檢核區 ========
                discrepancies = []
                empty_row = {c: "" for c in df_out_master.columns}
                row_act_d, row_req_d = empty_row.copy(), empty_row.copy()
                row_act_s, row_req_s = empty_row.copy(), empty_row.copy()
                row_diff = empty_row.copy()
                
                empty_row[df_out_master.columns[teacher_col_idx]] = "--- 系統自動檢核區 ---"
                row_act_d[df_out_master.columns[teacher_col_idx]] = "實際排入 (△)"
                row_req_d[df_out_master.columns[teacher_col_idx]] = "需求總數 (△)"
                row_act_s[df_out_master.columns[teacher_col_idx]] = "實際排入 (※)"
                row_req_s[df_out_master.columns[teacher_col_idx]] = "需求總數 (※)"
                row_diff[df_out_master.columns[teacher_col_idx]]  = "異常差額警示"

                for j in range(ai_periods):
                    col_name = df_out_master.columns[ai_period_cols[j]]
                    period_name = str(df_list_raw.iloc[header_row_idx, ai_period_cols[j]]).strip()
                    day_name = "Day1" if j in d1_idx else "Day2"
                    
                    act_d, req_d = actual_matrix['△'][j], req_matrix['△'][j]
                    act_s, req_s = actual_matrix['※'][j], req_matrix['※'][j]
                    
                    row_act_d[col_name], row_req_d[col_name] = act_d, req_d
                    row_act_s[col_name], row_req_s[col_name] = act_s, req_s
                    
                    diff_d, diff_s = act_d - req_d, act_s - req_s
                    diff_strs = []
                    
                    if diff_d != 0:
                        diff_strs.append(f"△{'+' if diff_d>0 else ''}{diff_d}")
                        discrepancies.append(f"【{day_name}】第 {period_name} 節 - △：需求 {req_d} 人，實際 {act_d} 人 (差額: {diff_d})")
                    if diff_s != 0:
                        diff_strs.append(f"※{'+' if diff_s>0 else ''}{diff_s}")
                        discrepancies.append(f"【{day_name}】第 {period_name} 節 - ※：需求 {req_s} 人，實際 {act_s} 人 (差額: {diff_s})")
                        
                    row_diff[col_name] = "、".join(diff_strs) if diff_strs else "正常吻合"

                empty_row_spacer = {c: "" for c in df_out_master.columns}
                summary_df = pd.DataFrame([empty_row, row_act_d, row_req_d, row_act_s, row_req_s, row_diff, empty_row_spacer])
                
                df_out_master = pd.concat([summary_df, df_out_master], ignore_index=True)

                # --- 3. 監考一覽表分配邏輯 ---
                with st.spinner("🎯 執行班級自動分配..."):
                    # 【記憶體防護】：加上 dropna
                    df_assign_calc = pd.read_excel(file_assign, header=None).dropna(how='all').fillna("")
                    raw_list = df_assign_calc.iloc[:, 0].astype(str).str.strip().tolist()
                    class_names_raw = [x for x in raw_list if x and not any(bad in x for bad in ["班級", "日期", "節次", "星期", "一覽表", "總表", "華南", "期中考", "註"])]
                    
                    norm_class_names = [normalize_cls(c) for c in class_names_raw]
                    assign_map = {name: idx for idx, name in enumerate(norm_class_names)}
                    
                    t2c_map = {}
                    for _, row in edited_bind_df.iterrows():
                        t_name = str(row['老師']).strip()
                        c_name = normalize_cls(row['班級']) 
                        if t_name and t_name != 'None' and c_name in assign_map:
                            t2c_map[t_name] = assign_map[c_name]
                    
                    assigned_matrix = np.empty((len(class_names_raw), ai_periods), dtype=object)
                    
                    for i_day, day_start in enumerate(day_starts):
                        day_end = day_starts[i_day+1] if i_day+1 < len(day_starts) else ai_periods
                        day_length = day_end - day_start
                        
                        j1 = day_start
                        proctors_j1 = [t for t in teachers if schedule_dict[t][j1] in ["△", "※"]]
                        random.shuffle(proctors_j1)
                        
                        rem_j1 = []
                        for p in proctors_j1:
                            if p in t2c_map and assigned_matrix[t2c_map[p], j1] is None:
                                assigned_matrix[t2c_map[p], j1] = p
                            else:
                                rem_j1.append(p)
                                
                        r_ptr = 0
                        for idx in range(len(class_names_raw)):
                            if assigned_matrix[idx, j1] is None and r_ptr < len(rem_j1):
                                assigned_matrix[idx, j1] = rem_j1[r_ptr]
                                r_ptr += 1
                        
                        if day_length > 1:
                            j2 = day_start + 1
                            proctors_j2 = [t for t in teachers if schedule_dict[t][j2] in ["△", "※"]]
                            bound = {}
                            
                            for idx in range(len(class_names_raw)):
                                p_prev = assigned_matrix[idx, j1]
                                if p_prev is not None and p_prev in schedule_dict:
                                    if schedule_dict[p_prev][j1] == "※" and schedule_dict[p_prev][j2] == "△":
                                        assigned_matrix[idx, j2] = p_prev
                                        bound[p_prev] = True
                            
                            rem = [p for p in proctors_j2 if p not in bound]
                            random.shuffle(rem)
                            
                            rem_after_bind = []
                            for p in rem:
                                if p in t2c_map and assigned_matrix[t2c_map[p], j2] is None:
                                    assigned_matrix[t2c_map[p], j2] = p
                                else:
                                    rem_after_bind.append(p)
                                    
                            r_idx = 0
                            for idx in range(len(class_names_raw)):
                                if assigned_matrix[idx, j2] is None and r_idx < len(rem_after_bind):
                                    assigned_matrix[idx, j2] = rem_after_bind[r_idx]; r_idx += 1

                            for offset in range(2, day_length):
                                curr_j = day_start + offset
                                proctors = [t for t in teachers if schedule_dict[t][curr_j] in ["△", "※"]]
                                random.shuffle(proctors)
                                
                                rem_curr = []
                                for p in proctors:
                                    if p in t2c_map and assigned_matrix[t2c_map[p], curr_j] is None:
                                        assigned_matrix[t2c_map[p], curr_j] = p
                                    else:
                                        rem_curr.append(p)
                                        
                                r_ptr = 0
                                for idx in range(len(class_names_raw)):
                                    if assigned_matrix[idx, curr_j] is None and r_ptr < len(rem_curr):
                                        assigned_matrix[idx, curr_j] = rem_curr[r_ptr]; r_ptr += 1

                    class_proctor_schedule = {} 
                    for r_idx, c_name in enumerate(class_names_raw):
                        class_proctor_schedule[normalize_cls(c_name)] = [assigned_matrix[r_idx, col] for col in range(ai_periods)]

                    wb_assign = openpyxl.load_workbook(file_assign)
                    ws_assign = wb_assign.active
                    
                    manual_proctors = {} 
                    first_class_row, class_col_idx = -1, 1
                    for r in range(1, 20):
                        for c in range(1, 5):
                            v = ws_assign.cell(row=r, column=c).value
                            if v and str(v).strip() in class_names_raw:
                                first_class_row, class_col_idx = r, c; break
                        if first_class_row != -1: break
                    
                    if first_class_row != -1:
                        target_cols = [class_col_idx + 1 + i for i in range(total_periods)]
                        
                        manual_assign_cols = []
                        ai_assign_cols = []
                        for i, c in enumerate(period_cols):
                            if c in manual_cols: manual_assign_cols.append(target_cols[i])
                            else: ai_assign_cols.append(target_cols[i])
                                
                        date_row = -1
                        for r in range(first_class_row - 1, max(0, first_class_row - 4), -1):
                            val = str(ws_assign.cell(row=r, column=target_cols[0]).value).strip()
                            if val != "" and not val.isdigit() and "期中" not in val and "華南" not in val:
                                date_row = r; break
                        
                        if date_row != -1:
                            if has_manual and d0_date:
                                for mc_a in manual_assign_cols:
                                    try: ws_assign.cell(row=date_row, column=mc_a).value = d0_date.strftime('%m月%d日')
                                    except AttributeError: pass
                            for j in range(ai_periods):
                                try: ws_assign.cell(row=date_row, column=ai_assign_cols[j]).value = get_ai_date_str(j, day_starts, ai_date_strs)
                                except AttributeError: pass
                        
                        for r in range(first_class_row, ws_assign.max_row + 1):
                            c_val = ws_assign.cell(row=r, column=class_col_idx).value
                            if c_val:
                                norm_c = normalize_cls(c_val)
                                
                                if has_manual:
                                    if norm_c not in manual_proctors: manual_proctors[norm_c] = {}
                                    for mc_idx, mc_a in enumerate(manual_assign_cols):
                                        orig_c = manual_cols[mc_idx]
                                        p_val = extract_period_num(str(df_list_raw.iloc[header_row_idx, orig_c]))
                                        val_m = ws_assign.cell(row=r, column=mc_a).value
                                        if val_m: manual_proctors[norm_c][p_val] = str(val_m).strip()
                                        
                                if norm_c in class_proctor_schedule:
                                    for j in range(ai_periods):
                                        ws_assign.cell(row=r, column=ai_assign_cols[j]).value = class_proctor_schedule[norm_c][j]
                    
                    out_assign = io.BytesIO()
                    wb_assign.save(out_assign)
                    assign_bytes = out_assign.getvalue()

                # --- 公布版套印 ---
                pub_bytes = None
                if file_pub:
                    with st.spinner("🖨️ 正在將資料無縫套印至公布版..."):
                        wb = openpyxl.load_workbook(file_pub)
                        ws = wb.active
                        h_row = -1; t_cols = []
                        for r in range(1, 16):
                            for c in range(1, 61):
                                val = ws.cell(row=r, column=c).value
                                if val and any(k in str(val) for k in ["教師", "姓名", "老師"]): h_row = r; t_cols.append(c)
                            if len(t_cols) > 0: break
                            
                        if h_row != -1:
                            for c in t_cols:
                                t_col_target = []
                                for scan_c in range(c + 1, c + 25):
                                    val = str(ws.cell(row=h_row, column=scan_c).value).strip()
                                    if any(k in val for k in ["教師", "姓名", "標號", "老師"]): break
                                    if extract_period_num(val) != -1: t_col_target.append(scan_c)
                                
                                if len(t_col_target) >= total_periods:
                                    pub_manual_cols = []
                                    pub_ai_cols = []
                                    for i, pc in enumerate(period_cols):
                                        if pc in manual_cols: pub_manual_cols.append(t_col_target[i])
                                        else: pub_ai_cols.append(t_col_target[i])
                                    
                                    if has_manual and d0_date:
                                        for pmc in pub_manual_cols:
                                            try: ws.cell(row=h_row-1, column=pmc).value = d0_date.strftime('%m月%d日')
                                            except AttributeError: pass
                                    
                                    for j in range(ai_periods):
                                        try: ws.cell(row=h_row-1, column=pub_ai_cols[j]).value = get_ai_date_str(j, day_starts, ai_date_strs)
                                        except AttributeError: pass
                                    
                                    for r in range(h_row+1, ws.max_row + 1):
                                        t_val = ws.cell(row=r, column=c).value
                                        if t_val:
                                            name = str(t_val).strip()
                                            if name in schedule_dict:
                                                for j in range(ai_periods):
                                                    ws.cell(row=r, column=pub_ai_cols[j]).value = schedule_dict[name][j]
                        out_pub = io.BytesIO()
                        wb.save(out_pub)
                        pub_bytes = out_pub.getvalue()

                # --- 標籤列印自動生成邏輯 ---
                label_bytes = None
                if file_course and file_label:
                    with st.spinner("🏷️ 正在合成試卷袋標籤..."):
                        course_dict = {}
                        xls_course = pd.ExcelFile(file_course)
                        for sheet in xls_course.sheet_names:
                            # 【記憶體防護】：加上 dropna
                            df_c = pd.read_excel(file_course, sheet_name=sheet).dropna(how='all').fillna("")
                            for r_idx, row in df_c.iterrows():
                                subj_raw = str(row.iloc[0]).strip()
                                if not subj_raw: continue
                                subj_norm = normalize_subject(subj_raw)
                                for c_idx in range(1, len(df_c.columns)):
                                    cls_raw = str(df_c.columns[c_idx]).strip()
                                    teacher = str(row.iloc[c_idx]).strip()
                                    if teacher and cls_raw:
                                        course_dict[(normalize_cls(cls_raw), subj_norm)] = teacher
                        
                        wb_label = openpyxl.load_workbook(file_label)
                        ws_label = wb_label.active
                        col_map = {}
                        header_row = 1
                        for r in range(1, 6):
                            for c in range(1, ws_label.max_column + 1):
                                val = str(ws_label.cell(row=r, column=c).value).strip()
                                if "班級" in val and '班級' not in col_map: col_map['班級'] = c
                                elif "科目" in val and '科目' not in col_map: col_map['科目'] = c
                                elif "日期" in val and '日期' not in col_map: col_map['日期'] = c
                                elif "序號" in val and '序號' not in col_map: col_map['序號'] = c
                                elif "任課" in val and '任課教師' not in col_map: col_map['任課教師'] = c
                                elif "監考" in val and '監考老師' not in col_map: col_map['監考老師'] = c
                            if '班級' in col_map and '監考老師' in col_map:
                                header_row = r; break

                        d1_ymd, d1_short, d1_slash = d1_date.strftime('%Y-%m-%d'), d1_date.strftime('%m-%d'), d1_date.strftime('%Y/%m/%d')
                        d2_ymd, d2_short, d2_slash = d2_date.strftime('%Y-%m-%d'), d2_date.strftime('%m-%d'), d2_date.strftime('%Y/%m/%d')
                        if has_manual and d0_date:
                            d0_ymd, d0_short, d0_slash = d0_date.strftime('%Y-%m-%d'), d0_date.strftime('%m-%d'), d0_date.strftime('%Y/%m/%d')

                        day_p_val_to_ai_col = {}
                        curr_day_idx = 0
                        for j in range(ai_periods):
                            if j in day_starts and j != 0: curr_day_idx += 1
                            day_p_val_to_ai_col[(curr_day_idx, ai_period_nums[j])] = j

                        for r in range(header_row + 1, ws_label.max_row + 1):
                            if '班級' not in col_map: continue
                            cls_raw = ws_label.cell(row=r, column=col_map['班級']).value
                            if cls_raw is None or not str(cls_raw).strip(): continue
                            
                            subj_raw = ws_label.cell(row=r, column=col_map['科目']).value if '科目' in col_map else ""
                            
                            date_val = ws_label.cell(row=r, column=col_map['日期']).value if '日期' in col_map else ""
                            if isinstance(date_val, datetime): date_str = date_val.strftime('%Y-%m-%d')
                            else: date_str = str(date_val).split()[0].strip() if date_val is not None else ""
                                
                            seq_val = ws_label.cell(row=r, column=col_map['序號']).value if '序號' in col_map else ""
                            
                            cls = normalize_cls(cls_raw)
                            subj = normalize_subject(subj_raw)
                            
                            if '任課教師' in col_map:
                                teacher = course_dict.get((cls, subj), "")
                                if not teacher:
                                    for (c, s), t in course_dict.items():
                                        if c == cls and (subj in s or s in subj):
                                            teacher = t; break
                                if teacher: ws_label.cell(row=r, column=col_map['任課教師']).value = teacher
                            
                            try: p_val = int(float(str(seq_val).strip()))
                            except: p_val = -1
                            
                            if '監考老師' in col_map:
                                if has_manual and d0_date and any(d in date_str for d in [d0_ymd, d0_short, d0_slash]):
                                    if cls in manual_day0_proctors:
                                        ws_label.cell(row=r, column=col_map['監考老師']).value = manual_day0_proctors[cls]
                                
                                elif cls in class_proctor_schedule and p_val != -1:
                                    day_idx = -1
                                    if any(d in date_str for d in [d1_ymd, d1_short, d1_slash]): day_idx = 0
                                    elif any(d in date_str for d in [d2_ymd, d2_short, d2_slash]): day_idx = 1
                                    
                                    if day_idx != -1 and (day_idx, p_val) in day_p_val_to_ai_col:
                                        target_col = day_p_val_to_ai_col[(day_idx, p_val)]
                                        ws_label.cell(row=r, column=col_map['監考老師']).value = class_proctor_schedule[cls][target_col]

                        out_label = io.BytesIO()
                        wb_label.save(out_label)
                        label_bytes = out_label.getvalue()

                st.session_state['results_p2'] = {
                    'orig': to_excel_bytes(df_out_master, header_df),
                    'assign': assign_bytes,
                    'pub': pub_bytes,
                    'label': label_bytes,
                    'discrepancies': discrepancies
                }
                
                if not discrepancies:
                    st.balloons()
                    st.success("✅ 完美排班！所有節次的監考人數與「監考類型總數」100% 吻合！")
                else:
                    st.warning("⚠️ 檢核提示：因您設定了特定鎖定條件（如特定日期或時段），部分節次的排入人數與原始需求有落差，明細如下：")
                    for d in discrepancies:
                        st.write(f"- {d}")
                    st.info("💡 下載出來的「1. 監考總表.xlsx」檔案最上方，也有為您附上完整的對帳明細喔！")

            except Exception as e:
                st.error("🚨 **系統攔截到未預期的中斷！**")
                st.warning("請將下方的「工程診斷報告」截圖或複製給您的 AI 工程師：")
                st.code(traceback.format_exc(), language="python")

    # ==========================================
    # 5. 下載區 (供階段二使用)
    # ==========================================
    if st.session_state['results_p2']:
        st.divider()
        res = st.session_state['results_p2']
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.download_button("📥 1. 監考總表", res['orig'], "監考總表.xlsx", "application/vnd.ms-excel", use_container_width=True, key="dl_p2_1")
        with c2: st.download_button("📥 2. 監考一覽表(保留手排)", res['assign'], "監考一覽表_分配完成.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="dl_p2_2")
        with c3: 
            if res['pub']: st.download_button("📥 3. 公布版套印總表", res['pub'], "公布版總表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_p2_3")
        with c4:
            if res.get('label'): st.download_button("📥 4. 標籤列印(完美接合)", res['label'], "標籤列印_完整版.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="dl_p2_4")
