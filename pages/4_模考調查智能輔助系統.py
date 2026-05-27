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
st.title("📊 教務處-模擬考調查智能輔助系統 (動態工作表切換版)")
st.info("💡 試務組終極進化：已導入「AI彈性列高演算引擎」，無論班級人數多寡，系統將自動撐滿版面並保證絕對【一班一頁】不溢出！")

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
    st.subheader("🛠️ 製作公版模擬考意願調查表 (支援跨表自動套印與動態考科)")
    st.markdown("上傳學生名單與「多工作表預設檔」，普高每次不同的考科組合，系統都能讓您下拉選擇並動態生成專屬調查表！")
    
    col1_t1, col2_t1 = st.columns([1, 1], gap="large")
    
    with col1_t1:
        file_roster = st.file_uploader("📥 1. 上傳學生原始名條 (必填，支援多工作表)", type=['xlsx', 'xls', 'csv'], key="roster_uploader")
        
        selected_roster_sheet = 0
        if file_roster and not file_roster.name.endswith('.csv'):
            try:
                xls_roster = pd.ExcelFile(file_roster)
                sheet_names = xls_roster.sheet_names
                if len(sheet_names) > 1:
                    selected_roster_sheet = st.selectbox("📑 偵測到名條有多個工作表，請選擇欲處理的名單：", sheet_names)
                else:
                    selected_roster_sheet = sheet_names[0]
            except Exception as e:
                st.warning("無法解析名單工作表，將預設讀取第一個。")

        file_preset = st.file_uploader("📥 2. 上傳【多工作表預設考科檔】 (選填)", type=['xlsx', 'xls'], key="preset_uploader")
        
        selected_preset_sheet = None
        if file_preset:
            try:
                xls_preset = pd.ExcelFile(file_preset)
                p_sheet_names = xls_preset.sheet_names
                if len(p_sheet_names) > 1:
                    options = p_sheet_names[1:]
                    st.success("✅ 已偵測到多個考科費用分頁！")
                    selected_preset_sheet = st.selectbox("📝 請選擇本次對應的【模擬考考科費用表】：", options)
                else:
                    selected_preset_sheet = p_sheet_names[0]
            except Exception as e:
                st.warning("無法解析設定檔工作表。")

        school_type = st.radio("🏫 選擇產出的學制類型：", ["技高 (統測群類)", "普高 (學測考科)"], horizontal=True)
        
    with col2_t1:
        default_title = "114學年度國立華南高商統測模擬考 報考類組調查表" if "技高" in school_type else "114學年度國立華南高商學測模擬考 報考考科調查表"
        template_name = st.text_input("🎯 擬定表單大標題", value=default_title)
        default_price = st.number_input("💰 預設單次費用 (無對照檔或查無資料時套用，可留 0)", min_value=0, max_value=2000, value=0, step=10)
        deadline_date = st.date_input("📅 調查表繳回截止日", value=datetime.today())
        deadline_str = f"{deadline_date.month}月{deadline_date.day}日"
        
    if st.button("🚀 生成一班一頁調查表", type="primary", use_container_width=True, key="btn_gen_template"):
        if not file_roster:
            st.error("🚨 請先上傳學生原始名條！")
        else:
            with st.spinner("正在智能合成調查表與動態考科套印中..."):
                try:
                    preset_mapping = {}
                    dynamic_target_mapping = [] 
                    fee_map = {}
                    
                    if file_preset:
                        xls = pd.ExcelFile(file_preset)
                        df_preset_dept = pd.read_excel(xls, sheet_name=0).fillna("")
                        
                        if selected_preset_sheet:
                            df_preset_fee = pd.read_excel(xls, sheet_name=selected_preset_sheet).fillna("")
                            
                            f_code = get_str_col(df_preset_fee, ['代碼', '類別', '群別', '類組', '代號'])
                            f_name = get_str_col(df_preset_fee, ['名稱', '考科', '組合', '類組名稱', '類別名稱'])
                            f_fee = get_str_col(df_preset_fee, ['費用', '金額', '單價', '單次費用'])
                            
                            for c_val, n_val, f_val in zip(f_code, f_name, f_fee):
                                c_str = str(c_val).strip().split('.')[0]
                                n_str = str(n_val).strip()
                                f_str = str(f_val).strip()
                                if c_str:
                                    fee_map[c_str] = f_str
                                    final_name = n_str if n_str else c_str
                                    if c_str not in [x[0] for x in dynamic_target_mapping]:
                                        dynamic_target_mapping.append((c_str, final_name, f_str))
                        else:
                            f_code = get_str_col(df_preset_dept, ['代碼', '類組', '類別'])
                            f_fee = get_str_col(df_preset_dept, ['費用', '金額', '單價'])
                            for c_val, f_val in zip(f_code, f_fee):
                                c_str = str(c_val).strip().split('.')[0]
                                f_str = str(f_val).strip()
                                if c_str and f_str: fee_map[c_str] = f_str

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

                    if file_roster.name.endswith('.csv'):
                        df_roster = pd.read_csv(file_roster).fillna("")
                    else:
                        df_roster = pd.read_excel(file_roster, sheet_name=selected_roster_sheet).fillna("")
                        
                    df_temp = pd.DataFrame()
                    df_temp['班級'] = get_str_col(df_roster, ['班級', '科別'])
                    df_temp['座號'] = get_str_col(df_roster, ['座號'])
                    df_temp['學號'] = get_str_col(df_roster, ['學號'])
                    df_temp['姓名'] = get_str_col(df_roster, ['姓名', '學生姓名'])
                    
                    df_temp['班級'] = df_temp['班級'].apply(clean_class_name)
                    df_temp = df_temp[(df_temp['班級'] != "") & (df_temp['姓名'] != "")].copy()
                    
                    df_temp['座號_Num'] = pd.to_numeric(df_temp['座號'], errors='coerce').fillna(999)
                    df_temp = df_temp.sort_values(by=['班級', '座號_Num']).drop(columns=['座號_Num'])
                    
                    df_temp['單次費用'] = default_price if default_price > 0 else ""
                    df_temp['報考類組'] = ""
                    df_temp['簽名'] = ""

                    for idx, row in df_temp.iterrows():
                        cls_name = str(row['班級'])
                        matched = preset_mapping.get(cls_name)
                        if not matched:
                            for k, v in preset_mapping.items():
                                if k in cls_name or cls_name.startswith(k):
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
                        
                        note_format_top = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'top': 2, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        note_format_middle = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        note_format_bottom = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bottom': 2, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        note_format_single = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'top': 2, 'bottom': 2, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        
                        red_alert_format = workbook.add_format({'font_color': '#D32F2F', 'bold': True, 'font_size': 9})
                        blue_alert_format = workbook.add_format({'font_color': '#1976D2', 'bold': True, 'font_size': 9})
                        
                        headers = ['班級', '座號', '學號', '姓名', '單次費用', '報考類組', '簽名']
                        
                        if "技高" in school_type:
                            target_mapping = [(k, v, fee_map.get(k, str(default_price)) if fee_map.get(k, "") != "" else str(default_price)) for k, v in STANDARD_MAPPING_VOC.items()]
                        else:
                            if dynamic_target_mapping:
                                target_mapping = dynamic_target_mapping
                            else:
                                target_mapping = [(k, v, fee_map.get(k, str(default_price)) if fee_map.get(k, "") != "" else str(default_price)) for k, v in STANDARD_MAPPING_GEN.items()]
                        
                        current_row = 0
                        page_breaks = []
                        unique_classes = df_temp['班級'].unique()
                        
                        is_gen_hs = "普高" in school_type
                        
                        for cls in unique_classes:
                            df_cls = df_temp[df_temp['班級'] == cls]
                            
                            merge_end_col = 6 if is_gen_hs else 9
                            rows_needed = len(df_cls) if is_gen_hs else max(len(df_cls), len(target_mapping))
                            
                            # 🚀 準備警語內容
                            if is_gen_hs and selected_preset_sheet:
                                if "高一" in selected_preset_sheet or "仿真" in selected_preset_sheet:
                                    memo_lines = [
                                        ["\n1.為讓同學了解學測考試時間及題型，將於二年級舉行第一次學測模擬考。"],
                                        ["2.請學藝股長於 ", red_alert_format, f"{deadline_str} 早上11點前", " 完成，此調查表交回教務處試務組。"],
                                        ["3.", red_alert_format, "報考類組請填代碼。"],
                                        ["4.模擬考費用將於新學期時9月初進行收取。\n"]
                                    ]
                                elif "第一" in selected_preset_sheet:
                                    memo_lines = [
                                        ["\n1.未參加暑期輔導的同學，不能參加第一次模擬考，", red_alert_format, "『報考類組』欄位請填不參加", "。"],
                                        ["2.請學藝股長於 ", red_alert_format, f"{deadline_str} 早上11點前", " 完成，此調查表交回教務處試務組。"],
                                        ["3.", red_alert_format, "報考類組請填代碼。"],
                                        ["4.模擬考費用將於調查表回收後，進行收取費用。\n"]
                                    ]
                                elif "第二" in selected_preset_sheet:
                                    memo_lines = [
                                        ["\n1.第二次模擬考，", blue_alert_format, "統一加考英聽", "。"],
                                        ["2.請學藝股長於 ", red_alert_format, f"{deadline_str} 早上11點前", " 完成，此調查表交回教務處試務組。"],
                                        ["3.", red_alert_format, "報考類組請填代碼。"],
                                        ["4.模擬考費用將於新學期時9月初進行收取費用。\n"]
                                    ]
                                elif "第三" in selected_preset_sheet or "第四" in selected_preset_sheet:
                                    memo_lines = [
                                        ["\n1.請學藝股長於 ", red_alert_format, f"{deadline_str} 早上11點前", " 完成，此張單子交回教務處『試務組』。"],
                                        ["2.", red_alert_format, "報考類組請填代碼。"],
                                        ["3.第三、四次費用將調查完畢後一起收取費用。\n"]
                                    ]
                                else:
                                    memo_lines = [
                                        ["\n1.請學藝股長協助調查考試類別，", red_alert_format, "如有更正請同學用紅筆更正並簽名", "，", blue_alert_format, "調查期間未到校者，簽名欄請空著不須代簽", "，", red_alert_format, f"此調查表請於 {deadline_str} 前交回教務處試務組。"],
                                        ["2.", red_alert_format, "報考類組請填代碼。"],
                                        ["3.上下學期總共參加5次模擬考，開學初進行收費相關事宜。\n"]
                                    ]
                            else:
                                memo_lines = [
                                    ["\n1.請學藝股長協助調查考試類別，", red_alert_format, "如有更正請同學用紅筆更正並簽名", "，", blue_alert_format, "調查期間未到校者，簽名欄請空著不須代簽", "，", red_alert_format, f"此調查表請於 {deadline_str} 前交回教務處試務組。"],
                                    ["2.", red_alert_format, "報考類組請填代碼。"],
                                    ["3.上下學期總共參加5次模擬考，開學初進行收費相關事宜。\n"]
                                ]

                            # 🚀 智能空間演算：預先計算警語高度
                            memo_heights = []
                            memo_h_sum = 0
                            for line_idx, rich_parts in enumerate(memo_lines):
                                text_length = sum(len(x) if type(x) == str else 0 for x in rich_parts)
                                if line_idx == 0:
                                    h = 38 if text_length > 60 else 30
                                elif line_idx == len(memo_lines) - 1:
                                    h = 28
                                else:
                                    h = 18
                                memo_heights.append(h)
                                memo_h_sum += h
                            
                            # 🚀 智能空間演算：加總固定高度，剩餘空間完美均分給學生資料
                            fixed_h = 25 + 20 + 10 + 35 + 15 + memo_h_sum
                            if is_gen_hs:
                                fixed_h += 20 + (len(target_mapping) * 18) + 10
                            
                            # A4 一頁可容納的安全總高度約 930 points
                            target_page_h = 930
                            calculated_data_h = (target_page_h - fixed_h) / max(1, rows_needed)
                            # 設定最大最小保護值 (下限16保證能看見，上限32避免太粗)
                            data_h = max(16, min(calculated_data_h, 32))
                            
                            # 開始寫入排版
                            worksheet.merge_range(current_row, 0, current_row, merge_end_col, template_name, title_format)
                            worksheet.set_row(current_row, 25)
                            current_row += 1
                            
                            for col_num, header in enumerate(headers):
                                worksheet.write(current_row, col_num, header, header_format)
                                
                            if not is_gen_hs:
                                worksheet.write(current_row, 8, "代碼", mapping_head_format)
                                worksheet.write(current_row, 9, "類別", mapping_head_format)
                                
                            worksheet.set_row(current_row, 20)
                            current_row += 1
                            
                            start_data_row = current_row
                            
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
                                        
                                if not is_gen_hs:
                                    if i < len(target_mapping):
                                        code, name, _ = target_mapping[i]
                                        worksheet.write(start_data_row + i, 8, code, mapping_data_format)
                                        worksheet.write(start_data_row + i, 9, name, mapping_data_format)
                                
                                # 套用彈性高度
                                worksheet.set_row(start_data_row + i, data_h)
                                
                            current_row = start_data_row + rows_needed
                            
                            worksheet.set_row(current_row, 10) 
                            current_row += 1
                            worksheet.merge_range(current_row, 0, current_row, merge_end_col, "導師確認簽章：________________________", signature_format)
                            worksheet.set_row(current_row, 35) 
                            current_row += 1
                            
                            worksheet.set_row(current_row, 15) 
                            current_row += 1
                            
                            if is_gen_hs:
                                worksheet.write(current_row, 0, "代碼", mapping_head_format)
                                worksheet.merge_range(current_row, 1, current_row, 4, "考科組合", mapping_head_format)
                                worksheet.merge_range(current_row, 5, current_row, 6, "單次費用", mapping_head_format)
                                worksheet.set_row(current_row, 20)
                                current_row += 1
                                
                                for code, name, fee in target_mapping:
                                    worksheet.write(current_row, 0, code, mapping_data_format)
                                    worksheet.merge_range(current_row, 1, current_row, 4, name, mapping_data_format)
                                    worksheet.merge_range(current_row, 5, current_row, 6, fee, mapping_data_format)
                                    worksheet.set_row(current_row, 18)
                                    current_row += 1
                                
                                worksheet.set_row(current_row, 10) 
                                current_row += 1

                            for line_idx, rich_parts in enumerate(memo_lines):
                                if len(memo_lines) == 1:
                                    fmt = note_format_single
                                elif line_idx == 0:
                                    fmt = note_format_top
                                elif line_idx == len(memo_lines) - 1:
                                    fmt = note_format_bottom
                                else:
                                    fmt = note_format_middle
                                    
                                worksheet.merge_range(current_row, 0, current_row, merge_end_col, "", fmt)
                                
                                has_format = any(type(x) != str for x in rich_parts)
                                if has_format:
                                    worksheet.write_rich_string(current_row, 0, *rich_parts, fmt)
                                else:
                                    worksheet.write(current_row, 0, rich_parts[0], fmt)
                                    
                                worksheet.set_row(current_row, memo_heights[line_idx])
                                current_row += 1

                            page_breaks.append(current_row)
                            
                        # 整體欄寬配置，G欄加寬
                        worksheet.set_column('A:B', 8)
                        worksheet.set_column('C:D', 10)
                        worksheet.set_column('E:F', 11) 
                        worksheet.set_column('G:G', 22) 
                        worksheet.set_column('H:H', 2) 
                        worksheet.set_column('I:I', 8) 
                        worksheet.set_column('J:J', 28) 
                        
                        if page_breaks:
                            worksheet.set_h_pagebreaks(page_breaks)
                            
                    st.session_state.template_excel_data = output_template.getvalue()
                    st.session_state.template_processed = True
                except Exception as e:
                    st.error("🚨 名條解析失敗，請檢查檔案格式。")
                    st.code(traceback.format_exc())

    if st.session_state.template_processed:
        school_prefix = "技高" if "技高" in school_type else "普高"
        st.success(f"🎉 {school_prefix}空白調查表生成完畢！彈性撐滿排版已完美運作。")
        st.download_button(
            label=f"📥 下載【{school_prefix} A4分頁版調查表】",
            data=st.session_state.template_excel_data,
            file_name=f"{template_name}_動態考科套印版_{datetime.now().strftime('%Y%m%d')}.xlsx",
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
            type=['xlsx', 'xls', 'csv'], 
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
                
                for r in range(len(df_preload)):
                    row_vals = [str(x).strip() for x in df_preload.iloc[r].tolist()]
                    
                    c_idx = -1
                    for i, val in enumerate(row_vals):
                        if val == '代碼':
                            c_idx = i
                            break
                    
                    if c_idx != -1:
                        n_idx = -1
                        for i in range(c_idx + 1, len(row_vals)):
                            if any(k in row_vals[i] for k in ['類別', '考科', '組合', '名稱', '類群']):
                                n_idx = i
                                break
                        
                        f_idx = -1
                        for i in range(c_idx + 1, len(row_vals)):
                            if '費用' in row_vals[i] or '金額' in row_vals[i] or '單價' in row_vals[i]:
                                f_idx = i
                                break
                        
                        if n_idx != -1:
                            for sub_r in range(r + 1, min(r + 40, len(df_preload))):
                                sub_row = [str(x).strip() for x in df_preload.iloc[sub_r].tolist()]
                                if len(sub_row) <= max(c_idx, n_idx): continue
                                
                                cv = sub_row[c_idx].split('.')[0]
                                nv = sub_row[n_idx]
                                
                                if any(k in cv for k in ['導師', '1.', '2.', '3.', '4.', '備註', '說明', '小計', '總計']):
                                    break
                                    
                                if cv and nv and cv != 'nan' and nv != 'nan':
                                    preload_mapping[cv] = nv
                                    if f_idx != -1 and len(sub_row) > f_idx:
                                        fv = sub_row[f_idx]
                                        if fv and fv != 'nan':
                                            try:
                                                extracted_fees[nv] = int(float(fv))
                                            except:
                                                pass

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
                    
                    raw_cat_series = get_str_col(df_data_preload, ['報考', '類群', '科目', '組別', '類組'])
                    raw_name_series = get_str_col(df_data_preload, ['姓名', '學生姓名'])
                    
                    unique_cats = set()
                    for cat_val, name_val in zip(raw_cat_series, raw_name_series):
                        if str(name_val).strip() == "" or str(name_val).strip() == "nan": 
                            continue
                            
                        cv = str(cat_val).strip().split('.')[0]
                        cat_name = ""
                        
                        if cv in preload_mapping:
                            cat_name = preload_mapping[cv]
                        elif cv and cv not in ["", "報考類組", "不升學", "休學", "重讀", "長期未到校", "否", "nan"] and not str(cat_val).startswith('*'):
                            cat_name = str(cat_val).strip()
                            
                        if cat_name:
                            unique_cats.add(cat_name)
                                
                    detected_categories = sorted(list(unique_cats))
            except Exception as e:
                st.error(f"預讀取檔案進行群別與費用分析時發生錯誤: {e}")

        st.markdown("📝 **列印優化說明**：")
        st.success("已擴充「學號」與「簽名」欄位！啟動 A4 彈性空間演算引擎，完美塞進一頁 A4 之中！")

    with col2:
        st.subheader("⚙️ 2. 收費檢核與測驗設定")
        
        school_type_p2 = st.radio("🏫 選擇本表單學制類型：", ["技高 (全學年5次合併收費)", "普高 (依次數彈性收費)"], horizontal=True, key="school_type_p2")
        
        if "普高" in school_type_p2:
            fee_mode = st.radio("🔄 普高本次收費模式：", ["收 1 次費用 (如：第一、二次模考)", "收 2 次費用 (如：第三、四次合併)"], horizontal=True)
            fee_multiplier = 1 if "1 次" in fee_mode else 2
            mock_name_default = "114學年度普高模擬考"
        else:
            fee_multiplier = 5
            mock_name_default = "114學年度高職模擬考(全學年共5次)"
            
        mock_name_p2 = st.text_input("🎯 產出報表標題", value=mock_name_default)
        base_fee_p2 = st.number_input("💰 預設單次報名費", min_value=0, max_value=5000, value=160, step=10)
        deadline_date_p2 = st.date_input("📅 收費明細表繳回截止日", value=datetime.today(), key="deadline_p2")
        deadline_str_p2 = f"{deadline_date_p2.month}月{deadline_date_p2.day}日"
        
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
                        df_raw_full = pd.read_csv(file_survey, header=None).fillna("")
                    else:
                        df_raw_full = pd.read_excel(file_survey, header=None).fillna("")

                    mapping_dict = {}
                    for r in range(len(df_raw_full)):
                        row_vals = [str(x).strip() for x in df_raw_full.iloc[r].tolist()]
                        c_idx = -1
                        for i, val in enumerate(row_vals):
                            if val == '代碼': c_idx = i; break
                        if c_idx != -1:
                            n_idx = -1
                            for i in range(c_idx + 1, len(row_vals)):
                                if any(k in row_vals[i] for k in ['類別', '考科', '組合', '名稱', '類群']): n_idx = i; break
                            if n_idx != -1:
                                for sub_r in range(r + 1, min(r + 40, len(df_raw_full))):
                                    sub_row = [str(x).strip() for x in df_raw_full.iloc[sub_r].tolist()]
                                    if len(sub_row) <= max(c_idx, n_idx): continue
                                    cv = sub_row[c_idx].split('.')[0]
                                    nv = sub_row[n_idx]
                                    if any(k in cv for k in ['導師', '1.', '2.', '3.', '4.', '備註', '小計', '總計']): break
                                    if cv and nv and cv != 'nan' and nv != 'nan':
                                        mapping_dict[cv] = nv

                    header_row_idx = None
                    for r in range(min(15, len(df_raw_full))):
                        row_vals = [str(x).strip() for x in df_raw_full.iloc[r].tolist()]
                        if '班級' in row_vals or '座號' in row_vals or '姓名' in row_vals:
                            header_row_idx = r
                            break
                    
                    if header_row_idx is not None:
                        new_columns = [str(x).strip() for x in df_raw_full.iloc[header_row_idx].tolist()]
                        df_raw = df_raw_full.iloc[header_row_idx+1:].copy()
                        df_raw.columns = new_columns
                    else:
                        st.error("🚨 找不到包含『班級』或『座號』的表頭欄位！")
                        st.stop()

                    df_all = pd.DataFrame()
                    df_all['班級_Raw'] = get_str_col(df_raw, ['班級', '科別'])
                    df_all['座號'] = get_str_col(df_raw, ['座號'])
                    df_all['學號'] = get_str_col(df_raw, ['學號'])
                    df_all['姓名'] = get_str_col(df_raw, ['姓名', '學生姓名'])
                    df_all['原始報考'] = get_str_col(df_raw, ['報考', '類群', '科目', '組別', '類組'])
                    df_all['班級_Clean'] = df_all['班級_Raw'].apply(clean_class_name)
                    
                    def is_valid_student(row):
                        c = str(row['班級_Clean'])
                        n = str(row['姓名'])
                        if not c or not n or c == 'nan' or n == 'nan': return False
                        if c in ['代碼', '班級', '姓名']: return False
                        if any(c.startswith(k) for k in ['1.', '2.', '3.', '4.', '導師']): return False
                        if n in ['姓名', '考科組合', '類別']: return False
                        return True
                        
                    df_students_only = df_all[df_all.apply(is_valid_student, axis=1)].copy()

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
                    df_clean['總繳費金額'] = df_clean['單次應繳費用'] * fee_multiplier
                    
                    df_clean['座號_Num'] = pd.to_numeric(df_clean['座號'], errors='coerce').fillna(999)
                    df_clean = df_clean.sort_values(by=['班級', '座號_Num']).drop(columns=['座號_Num'])

                    df_details_raw = df_clean[['班級', '座號', '學號', '姓名', '報考類群', '單次應繳費用', '總繳費金額']].copy()
                    df_details_raw = df_details_raw.rename(columns={'單次應繳費用': '單次費用', '總繳費金額': f'應繳費用({fee_multiplier}次)'})
                    
                    df_class_summary = df_clean.groupby('班級').agg(
                        報考人數=('姓名', 'count'),
                        應收總金額=('總繳費金額', 'sum')
                    ).reset_index()
                    
                    total_row = pd.DataFrame({
                        '班級': ['總計 (Total)'],
                        '報考人數': [df_class_summary['報考人數'].sum()],
                        '應收總金額': [df_class_summary['應收總金額'].sum()]
                    })
                    df_class_summary = pd.concat([df_class_summary, total_row], ignore_index=True)
                    df_class_summary = df_class_summary.rename(columns={'應收總金額': f'{fee_multiplier}次應收總金額'})

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
                        
                        sheet1_name = f'1_班級{fee_multiplier}次收費總表'
                        sheet2_name = f'2_各班繳費明細({fee_multiplier}次總額)'
                        
                        df_class_summary.to_excel(writer, index=False, sheet_name=sheet1_name)
                        workbook  = writer.book
                        worksheet1 = writer.sheets[sheet1_name]
                        worksheet1.set_column('A:F', 24) 
                        
                        start_row = len(df_class_summary) + 4
                        bold_format = workbook.add_format({'bold': True, 'font_color': '#D32F2F', 'font_size': 12})
                        worksheet1.write(start_row - 2, 0, "🔍 各班級未報考原因/狀態交叉檢核表", bold_format)
                        df_reason_by_class.to_excel(writer, index=False, sheet_name=sheet1_name, startrow=start_row)
                        
                        ws_details = workbook.add_worksheet(sheet2_name)
                        writer.sheets[sheet2_name] = ws_details 
                        
                        ws_details.set_paper(9)
                        ws_details.fit_to_pages(1, 0)
                        ws_details.center_horizontally()
                        ws_details.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4) 
                        
                        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#F2F2F2', 'border': 1})
                        header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter'})
                        data_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11}) 
                        cat_data_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9, 'shrink': True})
                        total_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#E2EFDA', 'align': 'center', 'valign': 'vcenter', 'font_size': 11})
                        memo_format = workbook.add_format({'font_size': 11, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FDFAD9'}) 
                        grand_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#FFF2CC', 'align': 'center', 'valign': 'vcenter', 'font_size': 12})
                        
                        note_format_top_p2 = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'top': 2, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        note_format_middle_p2 = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        note_format_bottom_p2 = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'bottom': 2, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        note_format_single_p2 = workbook.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True, 'top': 2, 'bottom': 2, 'left': 2, 'right': 2, 'border_color': '#D32F2F', 'indent': 1})
                        red_alert_format_p2 = workbook.add_format({'font_color': '#D32F2F', 'bold': True, 'font_size': 9})

                        headers = ['班級', '座號', '學號', '姓名', '報考類群', '單次費用', f'應繳費用({fee_multiplier}次)', '學生簽名']
                        
                        current_row = 0
                        page_breaks = []
                        unique_classes = df_details_raw['班級'].unique()
                        
                        for cls_name in unique_classes:
                            df_cls = df_details_raw[df_details_raw['班級'] == cls_name]
                            cls_count = len(df_cls)
                            
                            if "普高" in school_type_p2:
                                memo_lines_p2 = [
                                    ["\n1.請學藝股長於 ", red_alert_format_p2, f"{deadline_str_p2} 完成收費！", "費用請繳至教務處試務組。\n"]
                                ]
                            else:
                                memo_lines_p2 = [
                                    ["\n1.上下學期總共參加5次模擬考。"],
                                    ["2.開學初進行收費相關事宜。"],
                                    ["3.請學藝股長於 ", red_alert_format_p2, f"{deadline_str_p2} 完成收費！", "費用請繳至總務處出納組，此張單子請繳回到教務處試務組。\n"]
                                ]

                            # 🚀 智能空間演算：預先計算警語高度
                            memo_heights_p2 = []
                            memo_h_sum_p2 = 0
                            for line_idx, rich_parts in enumerate(memo_lines_p2):
                                text_length = sum(len(x) if type(x) == str else 0 for x in rich_parts)
                                if line_idx == 0:
                                    h = 36 if text_length > 60 else 30
                                elif line_idx == len(memo_lines_p2) - 1:
                                    h = 28
                                else:
                                    h = 18
                                memo_heights_p2.append(h)
                                memo_h_sum_p2 += h
                            
                            cls_fee_info = df_cls[['報考類群', '單次費用']].drop_duplicates().sort_values('報考類群')
                            
                            # 🚀 智能空間演算：計算階段二的固定高度 (不包含學生列表)
                            fixed_h_p2 = 24 + 18 + 26 + 18 + (len(cls_fee_info) * 16) + 15 + 35 + 15 + memo_h_sum_p2
                            
                            # 🚀 動態求出能最大化塞滿版面的學生列高 (總容忍高 930pt)
                            target_page_h_p2 = 930
                            calculated_data_h_p2 = (target_page_h_p2 - fixed_h_p2) / max(1, cls_count)
                            # 限制：最小 16 (大班微縮)，最大 32 (小班不至於太突兀)
                            data_h_p2 = max(16, min(calculated_data_h_p2, 32))

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
                                ws_details.write(current_row, 4, str(row['報考類群']), cat_data_format)
                                ws_details.write(current_row, 5, row['單次費用'], data_format)
                                ws_details.write(current_row, 6, row[f'應繳費用({fee_multiplier}次)'], data_format)
                                ws_details.write(current_row, 7, '', data_format) 
                                # 套用動態撐滿的彈性高度
                                ws_details.set_row(current_row, data_h_p2)
                                current_row += 1
                                
                            cls_total_amt = df_cls[f'應繳費用({fee_multiplier}次)'].sum()
                            
                            ws_details.write(current_row, 0, f'{cls_name} 小計', total_format)
                            ws_details.write(current_row, 1, f'共 {cls_count} 人', total_format)
                            ws_details.write(current_row, 2, '', total_format)
                            ws_details.write(current_row, 3, '', total_format)
                            ws_details.write(current_row, 4, '', total_format)
                            ws_details.write(current_row, 5, '', total_format)
                            ws_details.write(current_row, 6, cls_total_amt, total_format)
                            ws_details.write(current_row, 7, '', total_format) 
                            ws_details.set_row(current_row, 26) 
                            current_row += 1
                            
                            memo_title = f"💡 【本班單次報名費參考】 (※ 應繳總額 = 單次費用 × {fee_multiplier}次)："
                            ws_details.merge_range(current_row, 0, current_row, len(headers)-1, memo_title, memo_format)
                            ws_details.set_row(current_row, 18)
                            current_row += 1
                            
                            for _, r in cls_fee_info.iterrows():
                                bullet_text = f"      ▪ {r['報考類群']}：單次 {r['單次費用']} 元"
                                ws_details.merge_range(current_row, 0, current_row, len(headers)-1, bullet_text, memo_format)
                                ws_details.set_row(current_row, 16)
                                current_row += 1
                                
                            ws_details.set_row(current_row, 15) 
                            current_row += 1
                            
                            signature_format_p2 = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'right', 'valign': 'vcenter'})
                            ws_details.merge_range(current_row, 0, current_row, len(headers)-1, "導師確認簽章：________________________", signature_format_p2)
                            ws_details.set_row(current_row, 35) 
                            current_row += 1

                            ws_details.set_row(current_row, 15) 
                            current_row += 1

                            for line_idx, rich_parts in enumerate(memo_lines_p2):
                                if len(memo_lines_p2) == 1:
                                    fmt = note_format_single_p2
                                elif line_idx == 0:
                                    fmt = note_format_top_p2
                                elif line_idx == len(memo_lines_p2) - 1:
                                    fmt = note_format_bottom_p2
                                else:
                                    fmt = note_format_middle_p2
                                    
                                ws_details.merge_range(current_row, 0, current_row, len(headers)-1, "", fmt)
                                
                                has_format = any(type(x) != str for x in rich_parts)
                                if has_format:
                                    ws_details.write_rich_string(current_row, 0, *rich_parts, fmt)
                                else:
                                    ws_details.write(current_row, 0, rich_parts[0], fmt)
                                    
                                worksheet.set_row(current_row, memo_heights_p2[line_idx])
                                current_row += 1
                            
                            page_breaks.append(current_row) 
                            
                        ws_details.set_column('A:A', 8)  
                        ws_details.set_column('B:B', 6)  
                        ws_details.set_column('C:C', 10) 
                        ws_details.set_column('D:D', 10) 
                        ws_details.set_column('E:E', 22) 
                        ws_details.set_column('F:F', 10) 
                        ws_details.set_column('G:G', 12) 
                        ws_details.set_column('H:H', 24) # 大幅加寬「學生簽名」欄
                        
                        ws_details.write(current_row, 0, '全校總計 (Grand Total)', grand_format)
                        ws_details.write(current_row, 1, '', grand_format)
                        ws_details.write(current_row, 2, '', grand_format)
                        ws_details.write(current_row, 3, '', grand_format)
                        ws_details.write(current_row, 4, '', grand_format)
                        ws_details.write(current_row, 5, '', grand_format)
                        ws_details.write(current_row, 6, df_details_raw[f'應繳費用({fee_multiplier}次)'].sum(), grand_format)
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
