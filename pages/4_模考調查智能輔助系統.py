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
        else:
            with st.spinner("正在智能合成調查表與跨表自動套印中..."):
                try:
                    # 讀取雙工作表預設檔
                    preset_mapping = {}
                    if file_preset:
                        xls = pd.ExcelFile(file_preset)
                        df_preset_dept = pd.read_excel(xls, sheet_name=0).fillna("")
                        
                        # 讀取工作表 2 (費用字典)
                        fee_map = {}
                        if len(xls.sheet_names) > 1:
                            df_preset_fee = pd.read_excel(xls, sheet_name=1).fillna("")
                            f_code = get_str_col(df_preset_fee, ['代碼', '類組', '類別', '群別'])
                            f_fee = get_str_col(df_preset_fee, ['費用', '金額', '單價', '單次費用'])
                            for c_val, f_val in zip(f_code, f_fee):
                                c_str = str(c_val).strip().split('.')[0]
                                f_str = str(f_val).strip()
                                if c_str and f_str:
                                    fee_map[c_str] = f_str
                                    fee_map[str(c_val).strip()] = f_str
                        else:
                            # 備用防禦：如果老師只傳單一工作表，嘗試在第一頁找費用
                            f_code = get_str_col(df_preset_dept, ['代碼', '類組', '類別'])
                            f_fee = get_str_col(df_preset_dept, ['費用', '金額', '單價'])
                            for c_val, f_val in zip(f_code, f_fee):
                                c_str = str(c_val).strip().split('.')[0]
                                f_str = str(f_val).strip()
                                if c_str and f_str: fee_map[c_str] = f_str

                        # 讀取工作表 1 (科別對應) 並整合費用
                        s_key = get_str_col(df_preset_dept, ['科別', '班級', '字首'])
                        s_code = get_str_col(df_preset_dept, ['代碼', '類組', '預設'])
                        
                        for k, c in zip(s_key, s_code):
                            k_str = str(k).strip()
                            c_str = str(c).strip().split('.')[0]
                            if k_str:
                                preset_mapping[k_str] = {
                                    'code': c_str,
                                    'fee': fee_map.get(c_str, "")
                                }

                    # 讀取名條
                    if file_roster.name.endswith('.csv'):
                        df_roster = pd.read_csv(file_roster).fillna("")
                    else:
                        df_roster = pd.read_excel(file_roster).fillna("")
                        
                    df_temp = pd.DataFrame()
                    df_temp['班級'] = get_str_col(df_roster, ['班級', '科別'])
                    df_temp['座號'] = get_str_col(df_roster, ['座號'])
                    df_temp['學號'] = get_str_col(df_roster, ['學號'])
                    df_temp['姓名'] = get_str_col(df_roster, ['姓名', '學生姓名'])
                    
                    df_temp['班級'] = df_temp['班級'].apply(clean_class_name)
                    df_temp = df_temp[(df_temp['班級'] != "") & (df_temp['姓名'] != "")].copy()
                    
                    df_temp['座號_Num'] = pd.to_numeric(df_temp['座號'], errors='coerce').fillna(999)
                    df_temp = df_temp.sort_values(by=['班級', '座號_Num']).drop(columns=['座號_Num'])
                    
                    # 智慧自動套印
                    df_temp['單次費用'] = default_price if default_price > 0 else ""
                    df_temp['報考類組'] = ""
                    df_temp['簽名'] = ""

                    for idx, row in df_temp.iterrows():
                        cls_name = str(row['班級'])
                        matched = preset_mapping.get(cls_name)
                        if not matched:
                            # 支援模糊比對 (如科別寫「商」，可比對「商三1」)
                            for k, v in preset_mapping.items():
                                if cls_name.startswith(k):
                                    matched = v
                                    break
                        if matched:
                            if matched['code']: df_temp.at[idx, '報考類組'] = matched['code']
                            if matched['fee']: df_temp.at[idx, '單次費用'] = matched['fee']

                    output_template = io.BytesIO()
                    with pd.ExcelWriter(output_template, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet('調查表')
                        
                        worksheet.set_paper(9) # A4
                        worksheet.fit_to_pages(1, 0)
                        worksheet.center_horizontally()
                        worksheet.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
                        
                        title_format = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
                        header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter'})
                        data_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                        mapping_head_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'valign': 'vcenter'})
                        mapping_data_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                        signature_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'right', 'valign': 'vcenter'})
                        
                        headers = ['班級', '座號', '學號', '姓名', '單次費用', '報考類組', '簽名']
                        target_mapping = list((STANDARD_MAPPING_VOC if "技高" in school_type else STANDARD_MAPPING_GEN).items())
                        
                        current_row = 0
                        page_breaks = []
                        unique_classes = df_temp['班級'].unique()
                        
                        for cls in unique_classes:
                            df_cls = df_temp[df_temp['班級'] == cls]
                            
                            worksheet.merge_range(current_row, 0, current_row, 6, template_name, title_format)
                            worksheet.set_row(current_row, 25)
                            current_row += 1
                            
                            for col_num, header in enumerate(headers):
                                worksheet.write(current_row, col_num, header, header_format)
                            worksheet.write(current_row, 8, "代碼", mapping_head_format)
                            worksheet.write(current_row, 9, "類別", mapping_head_format)
                            worksheet.set_row(current_row, 20)
                            current_row += 1
                            
                            start_data_row = current_row
                            rows_needed = max(len(df_cls), len(target_mapping))
                            
                            for i in range(rows_needed):
                                if i < len(df_cls):
                                    row_data = df_cls.iloc[i]
                                    worksheet.write(start_data_row + i, 0, str(row_data['班級']), data_format)
                                    worksheet.write(start_data_row + i, 1, str(row_data['座號']), data_format)
                                    worksheet.write(start_data_row + i, 2, str(row_data['學號']), data_format)
                                    worksheet.write(start_data_row + i, 3, str(row_data['姓名']), data_format)
                                    worksheet.write(start_data_row + i, 4, str(row_data['單次費用']), data_format)
                                    worksheet.write(start_data_row + i, 5, str(row_data['報考類組']), data_format)
                                    worksheet.write(start_data_row + i, 6, "", data_format)
                                else:
                                    for c in range(7):
                                        worksheet.write(start_data_row + i, c, "", data_format)
                                        
                                if i < len(target_mapping):
                                    code, name = target_mapping[i]
                                    worksheet.write(start_data_row + i, 8, code, mapping_data_format)
                                    worksheet.write(start_data_row + i, 9, name, mapping_data_format)
                                
                                worksheet.set_row(start_data_row + i, 18)
                                
                            current_row = start_data_row + rows_needed
                            
                            worksheet.set_row(current_row, 10) 
                            current_row += 1
                            worksheet.merge_range(current_row, 0, current_row, 6, "導師確認簽章：________________________", signature_format)
                            worksheet.set_row(current_row, 35) 
                            current_row += 1
                            
                            page_breaks.append(current_row)
                            
                        worksheet.set_column('A:B', 8)
                        worksheet.set_column('C:D', 10)
                        worksheet.set_column('E:G', 12)
                        worksheet.set_column('H:H', 3) 
                        worksheet.set_column('I:I', 10) 
                        worksheet.set_column('J:J', 24) 
                        
                        if page_breaks:
                            worksheet.set_h_pagebreaks(page_breaks)
                            
                    st.session_state.template_excel_data = output_template.getvalue()
                    st.session_state.template_processed = True
                except Exception as e:
                    st.error("🚨 名條解析失敗，請檢查檔案格式。")
                    st.code(traceback.format_exc())

    if st.session_state.template_processed:
        school_prefix = "技高" if "技高" in school_type else "普高"
        st.success(f"🎉 {school_prefix}空白調查表生成完畢！自動套印功能已完美執行。")
        st.download_button(
            label=f"📥 下載【{school_prefix} A4分頁版調查表】",
            data=st.session_state.template_excel_data,
            file_name=f"{template_name}_跨表套印範本_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

# ---------------------------------------------------------
# 【階段二：結算與試務報表】
# ---------------------------------------------------------
with tab2:
    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        st.subheader("📂 1. 上傳已填妥之調查表")
        file_survey = st.file_uploader(
            "📥 上傳回收之學生意願調查表 (普高/技高表單皆可直接辨識)", 
            type=['xlsx', 'csv'], 
            key=f"mock_f2_{st.session_state.mock_uploader_key}"
        )
        
        detected_categories = []
        preload_mapping = {}
        extracted_fees = {} 

        if file_survey:
            try:
                file_survey.seek(0)
                if file_survey.name.endswith('.csv'):
                    df_preload = pd.read_csv(file_survey, header=None).fillna("")
                else:
                    df_preload = pd.read_excel(file_survey, header=None).fillna("")
                    
                h_idx = None
                for r in range(min(15, len(df_preload))):
                    row_vals = [str(x).strip() for x in df_preload.iloc[r].tolist()]
                    if '班級' in row_vals or '座號' in row_vals or '姓名' in row_vals:
                        h_idx = r
                        break
                        
                if h_idx is not None:
                    df_headers = [str(x).strip() for x in df_preload.iloc[h_idx].tolist()]
                    df_data_preload = df_preload.iloc[h_idx+1:].copy()
                    df_data_preload.columns = df_headers
                    
                    c_col, n_col = None, None
                    for i, col in enumerate(df_data_preload.columns):
                        col_name = str(col).strip()
                        if col_name == '代碼' and i > 4: c_col = i
                        if col_name in ['類別', '類群', '名稱', '科別', '群別', '報考類組'] and i > 4: n_col = i
                        
                    if c_col is None or n_col is None:
                        for i, col in enumerate(df_data_preload.columns):
                            col_str = str(col).strip()
                            if '代碼' in col_str and i > 4: c_col = i
                            if any(k in col_str for k in ['類別', '類群', '群', '類']) and i > 4: n_col = i
                            
                    if c_col is not None and n_col is not None:
                        for r in range(len(df_data_preload)):
                            cv = str(df_data_preload.iloc[r, c_col]).strip().split('.')[0]
                            nv = str(df_data_preload.iloc[r, n_col]).strip()
                            if cv and nv and cv != 'nan' and nv != 'nan' and cv.isdigit():
                                preload_mapping[cv] = nv
                                
                    raw_cat_series = get_str_col(df_data_preload, ['報考', '類群', '科目', '組別', '類組'])
                    raw_fee_series = get_str_col(df_data_preload, ['單次費用', '費用', '單價', '金額', '報名費'])
                    
                    unique_cats = set()
                    for cat_val, fee_val in zip(raw_cat_series, raw_fee_series):
                        cv = str(cat_val).strip().split('.')[0]
                        cat_name = ""
                        if cv in preload_mapping:
                            cat_name = preload_mapping[cv]
                        elif cv and cv not in ["", "報考類組", "不升學", "休學", "重讀", "長期未到校", "否", "nan"] and not str(cat_val).startswith('*'):
                            cat_name = str(cat_val).strip()
                            
                        if cat_name:
                            unique_cats.add(cat_name)
                            try:
                                fee_num = int(float(str(fee_val).strip()))
                                if cat_name not in extracted_fees or extracted_fees[cat_name] == 0:
                                    extracted_fees[cat_name] = fee_num
                            except:
                                pass
                                
                    detected_categories = sorted(list(unique_cats))
            except Exception as e:
                st.error(f"預讀取檔案進行群別與費用分析時發生錯誤: {e}")

        st.markdown("📝 **列印優化說明**：")
        st.success("已擴充「學號」欄位！啟動 A4 極限微調排版，確保 40 人以上的大班級依然能完美塞進一頁 A4 之中！")

    with col2:
        st.subheader("⚙️ 2. 收費檢核與測驗設定")
        mock_name_p2 = st.text_input("🎯 產出報表標題", value="114學年度高職模擬考(全學年共5次)")
        base_fee_p2 = st.number_input("💰 預設單次報名費", min_value=0, max_value=5000, value=160, step=10)
        
        special_fee_dict = {}
        if file_survey and detected_categories:
            st.markdown("### 💰 各類群單次費用檢核表")
            st.caption("💡 系統已自動將您上傳表單中的金額帶入，請檢核。")
            
            default_fees = []
            for cat in detected_categories:
                default_fees.append(extracted_fees.get(cat, base_fee_p2)) 
                
            fee_df = pd.DataFrame({
                '報考類群': detected_categories,
                '單次費用 (元)': default_fees
            })
            
            edited_fee_df = st.data_editor(
                fee_df,
                column_config={
                    '報考類群': st.column_config.TextColumn("報考類群", disabled=True),
                    '單次費用 (元)': st.column_config.NumberColumn("單次費用 (元)", min_value=0, max_value=2000, step=10)
                },
                disabled=["報考類群"],
                hide_index=True,
                use_container_width=True,
                key="fee_editor"
            )
            special_fee_dict = dict(zip(edited_fee_df['報考類群'], edited_fee_df['單次費用 (元)']))

        st.write("")
        if st.button("🗑️ 清除第二階段設定", use_container_width=True):
            st.session_state.mock_processed = False
            st.session_state.mock_excel_data = None
            st.session_state.mock_uploader_key += 1
            st.rerun()

    # ==========================================
    # 核心排程與生成演算法
    # ==========================================
    st.divider()

    if st.button("🚀 啟動智能分流與結算", type="primary", use_container_width=True, key="btn_run_phase2"):
        if not file_survey:
            st.error("🚨 老師，請記得上傳【學生意願調查表】喔！")
        else:
            with st.spinner("系統正在高速進行結算與【含學號版 A4 壓縮排版】統計中..."):
                try:
                    file_survey.seek(0)
                    if file_survey.name.endswith('.csv'):
                        df_raw = pd.read_csv(file_survey, header=None).fillna("")
                    else:
                        df_raw = pd.read_excel(file_survey, header=None).fillna("")

                    header_row_idx = None
                    for r in range(min(15, len(df_raw))):
                        row_vals = [str(x).strip() for x in df_raw.iloc[r].tolist()]
                        if '班級' in row_vals or '座號' in row_vals or '姓名' in row_vals:
                            header_row_idx = r
                            break
                    
                    if header_row_idx is not None:
                        new_columns = [str(x).strip() for x in df_raw.iloc[header_row_idx].tolist()]
                        df_raw = df_raw.iloc[header_row_idx+1:].copy()
                        df_raw.columns = new_columns
                    else:
                        st.error("🚨 找不到包含『班級』或『座號』的表頭欄位！")
                        st.stop()

                    mapping_dict = {}
                    code_col_idx = None
                    name_col_idx = None
                    for i, col in enumerate(df_raw.columns):
                        col_name = str(col).strip()
                        if col_name == '代碼' and i > 4: code_col_idx = i
                        if col_name in ['類別', '類群', '名稱', '科別', '群別', '報考類組'] and i > 4: name_col_idx = i

                    if code_col_idx is None or name_col_idx is None:
                        for i, col in enumerate(df_raw.columns):
                            col_str = str(col).strip()
                            if '代碼' in col_str and i > 4: code_col_idx = i
                            if any(k in col_str for k in ['類別', '類群', '群', '類']) and i > 4: name_col_idx = i

                    if code_col_idx is not None and name_col_idx is not None:
                        for r in range(len(df_raw)):
                            cv = str(df_raw.iloc[r, code_col_idx]).strip().split('.')[0]
                            nv = str(df_raw.iloc[r, name_col_idx]).strip()
                            if cv and nv and cv != 'nan' and nv != 'nan' and cv.isdigit():
                                mapping_dict[cv] = nv

                    df_all = pd.DataFrame()
                    df_all['班級_Raw'] = get_str_col(df_raw, ['班級', '科別'])
                    df_all['座號'] = get_str_col(df_raw, ['座號'])
                    df_all['學號'] = get_str_col(df_raw, ['學號'])
                    df_all['姓名'] = get_str_col(df_raw, ['姓名', '學生姓名'])
                    df_all['原始報考'] = get_str_col(df_raw, ['報考', '類群', '科目', '組別', '類組'])
                    df_all['班級_Clean'] = df_all['班級_Raw'].apply(clean_class_name)
                    
                    df_students_only = df_all[
                        (df_all['班級_Clean'] != "") & 
                        (df_all['姓名'] != "") & 
                        (~df_all['班級_Clean'].str.contains('\*', na=False)) &
                        (df_all['姓名'].str.lower() != "姓名") &
                        (df_all['班級_Clean'].str.lower() != "班級")
                    ].copy()

                    def determine_status_or_cat(row):
                        raw = str(row['原始報考']).strip().split('.')[0]
                        if raw in mapping_dict:
                            return "VALID", mapping_dict[raw]
                        elif raw in ["不升學", "休學", "重讀", "長期未到校"]:
                            return "UNREPORTED", raw
                        elif raw in ["", "nan", "否", "0", "None"]:
                            return "UNREPORTED", "未填寫/未報考"
                        elif raw in ["報考類組", "類組", "科目"]:
                            return "IGNORE", ""
                        else:
                            return "UNREPORTED", str(row['原始報考']).strip()

                    statuses = []
                    cats = []
                    for _, row in df_students_only.iterrows():
                        st_type, cat_v = determine_status_or_cat(row)
                        statuses.append(st_type)
                        cats.append(cat_v)

                    df_students_only['狀態類型'] = statuses
                    df_students_only['解析結果'] = cats

                    df_unreported = df_students_only[df_students_only['狀態類型'] == "UNREPORTED"].copy()
                    if not df_unreported.empty:
                        df_reason_by_class = df_unreported.groupby(['班級_Clean', '解析結果']).size().unstack(fill_value=0)
                        df_reason_by_class['未報考小計'] = df_reason_by_class.sum(axis=1)
                        df_reason_by_class.index.name = '班級'
                        df_reason_by_class = df_reason_by_class.reset_index()
                        
                        total_dict = {'班級': '總計 (Total)'}
                        for col in df_reason_by_class.columns:
                            if col != '班級': total_dict[col] = df_reason_by_class[col].sum()
                        df_reason_by_class = pd.concat([df_reason_by_class, pd.DataFrame([total_dict])], ignore_index=True)
                    else:
                        df_reason_by_class = pd.DataFrame(columns=['班級', '未報考小計'])

                    df_clean = df_students_only[df_students_only['狀態類型'] == "VALID"].copy()
                    df_clean['報考類群'] = df_clean['解析結果']
                    df_clean['班級'] = df_clean['班級_Clean']

                    if df_clean.empty:
                        st.warning("⚠️ 攔截警告：經過分析後檔案內找不到有效報考資料。")
                        st.stop()

                    df_clean['單次應繳費用'] = df_clean['報考類群'].apply(lambda x: special_fee_dict.get(x, base_fee_p2))
                    df_clean['五次總繳費金額'] = df_clean['單次應繳費用'] * 5
                    
                    df_clean['座號_Num'] = pd.to_numeric(df_clean['座號'], errors='coerce').fillna(999)
                    df_clean = df_clean.sort_values(by=['班級', '座號_Num']).drop(columns=['座號_Num'])

                    df_details_raw = df_clean[['班級', '座號', '學號', '姓名', '報考類群', '單次應繳費用', '五次總繳費金額']].copy()
                    df_details_raw = df_details_raw.rename(columns={'單次應繳費用': '單次費用(參考)', '五次總繳費金額': '應繳總金額(5次)'})
                    
                    df_class_summary = df_clean.groupby('班級').agg(
                        報考人數=('姓名', 'count'),
                        五次應收總金額=('五次總繳費金額', 'sum')
                    ).reset_index()
                    
                    total_row = pd.DataFrame({
                        '班級': ['總計 (Total)'],
                        '報考人數': [df_class_summary['報考人數'].sum()],
                        '五次應收總金額': [df_class_summary['五次應收總金額'].sum()]
                    })
                    df_class_summary = pd.concat([df_class_summary, total_row], ignore_index=True)

                    df_publisher = df_clean.groupby('報考類群').agg(
                        需求卷數=('姓名', 'count')
                    ).reset_index()
                    df_publisher = df_publisher.sort_values('需求卷數', ascending=False)
                    total_pub_row = pd.DataFrame({'報考類群': ['總計 (Total)'], '需求卷數': [df_publisher['需求卷數'].sum()]})
                    df_publisher = pd.concat([df_publisher, total_pub_row], ignore_index=True)

                    # ==========================================
                    # 5. 【流式 Excel 直出封裝引擎 - 列印極限壓縮排版】
                    # ==========================================
                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                        
                        df_class_summary.to_excel(writer, index=False, sheet_name='1_總務處班級5次收費總表')
                        workbook  = writer.book
                        worksheet1 = writer.sheets['1_總務處班級5次收費總表']
                        worksheet1.set_column('A:F', 24) 
                        
                        start_row = len(df_class_summary) + 4
                        bold_format = workbook.add_format({'bold': True, 'font_color': '#D32F2F', 'font_size': 12})
                        worksheet1.write(start_row - 2, 0, "🔍 各班級未報考原因/狀態交叉檢核表", bold_format)
                        df_reason_by_class.to_excel(writer, index=False, sheet_name='1_總務處班級5次收費總表', startrow=start_row)
                        
                        ws_details = workbook.add_worksheet('2_各班繳費明細(5次總額)')
                        writer.sheets['2_各班繳費明細(5次總額)'] = ws_details 
                        
                        ws_details.set_paper(9)
                        ws_details.fit_to_pages(1, 0)
                        ws_details.center_horizontally()
                        ws_details.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4) 
                        
                        ws_details.set_column('A:A', 9)  
                        ws_details.set_column('B:B', 6)  
                        ws_details.set_column('C:C', 10) 
                        ws_details.set_column('D:D', 10) 
                        ws_details.set_column('E:E', 21) 
                        ws_details.set_column('F:F', 13) 
                        ws_details.set_column('G:G', 14) 
                        ws_details.set_column('H:H', 15) 
                        
                        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#F2F2F2', 'border': 1})
                        header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter'})
                        data_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11}) 
                        total_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#E2EFDA', 'align': 'center', 'valign': 'vcenter', 'font_size': 11})
                        memo_format = workbook.add_format({'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FDFAD9'}) 
                        grand_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
                        
                        headers = ['班級', '座號', '學號', '姓名', '報考類群', '單次費用(參考)', '應繳總金額(5次)', '學生簽名']
                        
                        current_row = 0
                        page_breaks = []
                        unique_classes = df_details_raw['班級'].unique()
                        
                        for cls_name in unique_classes:
                            df_cls = df_details_raw[df_details_raw['班級'] == cls_name]
                            
                            ws_details.merge_range(current_row, 0, current_row, len(headers)-1, f"🏫 國立華南高商 - {mock_name_p2}", title_format)
                            ws_details.set_row(current_row, 24) 
                            current_row += 1
                            
                            for col_num, header in enumerate(headers):
                                ws_details.write(current_row, col_num, header, header_format)
                            ws_details.set_row(current_row, 18)
                            current_row += 1
                            
                            for _, row in df_cls.iterrows():
                                ws_details.write(current_row, 0, str(row['班級']), data_format)
                                ws_details.write(current_row, 1, str(row['座號']), data_format)
                                ws_details.write(current_row, 2, str(row['學號']), data_format)
                                ws_details.write(current_row, 3, str(row['姓名']), data_format)
                                ws_details.write(current_row, 4, str(row['報考類群']), data_format)
                                ws_details.write(current_row, 5, row['單次費用(參考)'], data_format)
                                ws_details.write(current_row, 6, row['應繳總金額(5次)'], data_format)
                                ws_details.write(current_row, 7, '', data_format) 
                                ws_details.set_row(current_row, 16)
                                current_row += 1
                                
                            cls_total_amt = df_cls['應繳總金額(5次)'].sum()
                            cls_count = len(df_cls)
                            
                            ws_details.write(current_row, 0, f'{cls_name} 小計', total_format)
                            ws_details.write(current_row, 1, f'共 {cls_count} 人', total_format)
                            ws_details.write(current_row, 2, '', total_format)
                            ws_details.write(current_row, 3, '', total_format)
                            ws_details.write(current_row, 4, '', total_format)
                            ws_details.write(current_row, 5, '', total_format)
                            ws_details.write(current_row, 6, cls_total_amt, total_format)
                            ws_details.write(current_row, 7, '導師簽章：', total_format) 
                            ws_details.set_row(current_row, 26) 
                            current_row += 1
                            
                            cls_fee_info = df_cls[['報考類群', '單次費用(參考)']].drop_duplicates().sort_values('報考類群')
                            memo_title = "💡 【本班單次報名費參考】 (※ 應繳總額 = 單次費用 × 5次)："
                            ws_details.merge_range(current_row, 0, current_row, len(headers)-1, memo_title, memo_format)
                            ws_details.set_row(current_row, 18)
                            current_row += 1
                            
                            for _, r in cls_fee_info.iterrows():
                                bullet_text = f"      ▪ {r['報考類群']}：單次 {r['單次費用(參考)']} 元"
                                ws_details.merge_range(current_row, 0, current_row, len(headers)-1, bullet_text, memo_format)
                                ws_details.set_row(current_row, 16)
                                current_row += 1
                            
                            page_breaks.append(current_row) 
                        
                        ws_details.write(current_row, 0, '全校總計 (Grand Total)', grand_format)
                        ws_details.write(current_row, 1, '', grand_format)
                        ws_details.write(current_row, 2, '', grand_format)
                        ws_details.write(current_row, 3, '', grand_format)
                        ws_details.write(current_row, 4, '', grand_format)
                        ws_details.write(current_row, 5, '', grand_format)
                        ws_details.write(current_row, 6, df_details_raw['應繳總金額(5次)'].sum(), grand_format)
                        ws_details.write(current_row, 7, '', grand_format)
                        ws_details.set_row(current_row, 24)
                        
                        if page_breaks:
                            ws_details.set_h_pagebreaks(page_breaks)
                        
                        df_publisher.to_excel(writer, index=False, sheet_name='3_書商訂卷總表')
                        writer.sheets['3_書商訂卷總表'].set_column('A:B', 24)
                    
                    st.session_state.mock_excel_data = output_excel.getvalue()
                    st.session_state.mock_preview_df = df_class_summary
                    st.session_state.mock_processed = True

                except Exception as e:
                    st.error("🚨 發生未預期錯誤，請檢查檔案格式。")
                    with st.expander("點此查看詳細工程錯誤碼"):
                        st.code(traceback.format_exc())

    # ==========================================
    # 結果顯示與下載區
    # ==========================================
    if st.session_state.mock_processed:
        st.balloons()
        st.success("🎉 第二階段試務報表結算完成！")
        
        st.download_button(
            label="📥 點擊下載【模擬考收費與各班未報考人數交叉檢核總表】",
            data=st.session_state.mock_excel_data,
            file_name=f"{mock_name_p2}_結算總表_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
            key="btn_download_phase2"
        )

        st.write("👀 **第一個工作表 (包含班級收費) 預覽：**")
        st.dataframe(st.session_state.mock_preview_df, use_container_width=True)
