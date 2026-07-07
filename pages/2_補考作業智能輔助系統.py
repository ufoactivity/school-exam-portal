import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import traceback

# ==========================================
# 1. 網頁頁面配置
# ==========================================
st.set_page_config(page_title="補考自動化神器-頂規網頁版", page_icon="🏫", layout="wide")

st.title("📝 試務組-補考作業智能輔助系統")
st.info("💡 修正說明：更新報表三(考程匯整表)排序邏輯，優先依「場地」分區，次依「應到人數」排序，完美支援試卷袋標籤套印！")

# --- 初始化快取記憶體與清空鑰匙 ---
if 'results' not in st.session_state:
    st.session_state['results'] = None
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

# ==========================================
# 2. 輔助功能定義
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

def grade_to_chinese(text):
    t = str(text)
    if '一' in t: return '一'
    if '二' in t: return '二'
    if '三' in t: return '三'
    if '1' in t or '１' in t: return '一'
    if '2' in t or '２' in t: return '二'
    if '3' in t or '３' in t: return '三'
    return "未知"

def natural_sort_key(s):
    return tuple(int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s)))

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ==========================================
# 3. 介面佈局：功能選單
# ==========================================
st.divider()
col_files, col_opts = st.columns([1, 1], gap="large")

with col_files:
    st.subheader("📂 第一步：上傳原始資料")
    file_target = st.file_uploader("1️⃣ 補考名單.xlsx", type=['xlsx'], key=f"f1_{st.session_state['uploader_key']}")
    file_short = st.file_uploader("2️⃣ 科目簡稱.xlsx", type=['xlsx'], key=f"f2_{st.session_state['uploader_key']}")
    file_exam = st.file_uploader("3️⃣ 科目對照表.xlsx", type=['xlsx'], key=f"f3_{st.session_state['uploader_key']}")
    file_teacher = st.file_uploader("4️⃣ 監考教師及時間.xlsx", type=['xlsx'], key=f"f4_{st.session_state['uploader_key']}")

with col_opts:
    st.subheader("⚙️ 第二步：考場容量與分流設定")
    zhiyong_cap = st.radio("📍 致用樓4樓會議室 人數上限：", [136, 148], index=0, horizontal=True)
    st.write("") 
    separate_mode = st.toggle("🔥 開啟【多科與單科嚴格分流】功能", value=False)
    
    st.write("")
    st.write("")
    if st.button("🗑️ 清除舊資料 / 重新設定", use_container_width=True):
        st.session_state['results'] = None
        st.session_state['uploader_key'] += 1
        st.rerun()

# ==========================================
# 4. 執行與運算
# ==========================================
st.divider()

if st.button("🚀 開始智慧排考運算", type="primary", use_container_width=True):
    if not all([file_target, file_short, file_exam, file_teacher]):
        st.error("🚨 錯誤：請確認上方【4個檔案】皆已上傳完畢！")
    else:
        with st.spinner("系統正在執行智慧運算中..."):
            try:
                # 讀取 Excel
                df_short_map = pd.read_excel(file_short)
                df_exam_map = pd.read_excel(file_exam)
                df_target = pd.read_excel(file_target)
                df_teacher = pd.read_excel(file_teacher)

                grade_weight = {'一': 1, '二': 2, '三': 3}
                loc_weight = {'致用樓四樓會議室': 1, '圖書館三樓自修教室': 2, '電腦教室401': 3}
                current_targets = {"致用樓四樓會議室": zhiyong_cap, "圖書館三樓自修教室": 98, "電腦教室401": 37}

                # --- 階段一：基本資料處理 ---
                df_target['姓名'] = get_str_col(df_target, ['姓名', '學生姓名'])
                col_opencourse = get_str_col(df_target, ['開課班'])
                col_homeroom = get_str_col(df_target, ['班級', '原班級'])
                if col_opencourse.eq("").all(): col_opencourse = col_homeroom

                df_target['學號'] = get_str_col(df_target, ['學號'])
                df_target['座號'] = get_str_col(df_target, ['座號'])
                df_target['科目'] = get_str_col(df_target, ['科目', '考科'])
                df_target['班級'] = col_homeroom
                df_target['年級'] = df_target['班級'].apply(grade_to_chinese)

                col_a_s, col_b_s = df_short_map.columns[0], df_short_map.columns[1]
                short_dict = dict(zip(df_short_map[col_a_s].astype(str).str.strip(), df_short_map[col_b_s].astype(str).str.strip()))
                df_target['科目簡稱'] = df_target['科目'].map(short_dict).fillna("")
                
                ex_cls = get_str_col(df_exam_map, ['班級', '開課班'])
                ex_sub = get_str_col(df_exam_map, ['科目', '考科'])
                ex_pap = get_str_col(df_exam_map, ['試卷編號', '代碼'])
                if ex_pap.eq("").all() and df_exam_map.shape[1] > 7: ex_pap = df_exam_map.iloc[:, 7].astype(str).str.strip()
                ex_dict = dict(zip(ex_cls + ex_sub, ex_pap))
                
                df_target['試卷編號'] = (col_opencourse + df_target['科目']).map(ex_dict).fillna((col_homeroom + df_target['科目']).map(ex_dict)).fillna("")
                df_target['試卷編號'] = df_target['試卷編號'].apply(lambda x: str(x).replace('.0','') if str(x).endswith('.0') else str(x))

                # 場地分配
                df_temp = df_target.drop_duplicates(subset=['學號', '試卷編號'], keep='first')
                v_counts = df_temp[df_temp['試卷編號'] != ""].groupby('學號').size()
                df_target['科目數目'] = df_target['學號'].map(v_counts).fillna(0).astype(int)
                
                df_students = df_target.drop_duplicates(subset=['學號']).copy()
                df_students = df_students[df_students['科目數目'] > 0].sort_values(by=['年級', '科目數目', '班級', '座號'], ascending=[True, False, True, True])

                venue_map = {}
                for gr, group in df_students.groupby('年級'):
                    if separate_mode:
                        multi = group[group['科目數目'] >= 2]
                        single = group[group['科目數目'] == 1]
                        m_v = (['致用樓四樓會議室'] * zhiyong_cap + ['圖書館三樓自修教室'] * 98 + ['電腦教室401'] * 37)
                        m_ans = m_v[:len(multi)]
                        rem_l = max(0, 98 - m_ans.count('圖書館三樓自修教室'))
                        rem_c = max(0, 37 - m_ans.count('電腦教室401'))
                        s_v = (['圖書館三樓自修教室'] * rem_l + ['電腦教室401'] * rem_c)
                        s_ans = s_v[:len(single)]
                        for sid, v in zip(multi['學號'], m_ans): venue_map[sid] = v
                        for sid, v in zip(single['學號'], s_ans): venue_map[sid] = v
                    else:
                        vns = (['致用樓四樓會議室'] * zhiyong_cap + ['圖書館三樓自修教室'] * 98 + ['電腦教室401'] * 37)
                        for sid, v in zip(group['學號'], vns[:len(group)]): venue_map[sid] = v

                df_target['場地'] = df_target['學號'].map(venue_map).fillna("")
                
                df_teacher['監考教師'] = get_str_col(df_teacher, ['監考教師', '監考老師', '教師姓名', '老師'])
                df_teacher['場地'] = get_str_col(df_teacher, ['場地', '地點', '考場', '場點'])
                df_teacher['比對年級'] = get_str_col(df_teacher, ['監考年級', '年級']).apply(grade_to_chinese)
                t_map = df_teacher.drop_duplicates(subset=['比對年級']).set_index('比對年級')[get_str_col(df_teacher, ['時間']).name].to_dict()
                df_target['時間2'] = df_target['年級'].map(t_map).fillna("")

                # --- 階段二：報表二處理 (標籤) ---
                df_label = df_target.drop_duplicates(subset=['學號', '試卷編號'], keep='first').copy()
                df_label['單科標籤'] = df_label.apply(lambda r: f"{r['試卷編號']}{r['科目簡稱']}" if str(r['試卷編號']).strip() != "" else "", axis=1)
                
                group_cols = ['年級', '班級', '座號', '姓名', '科目數目', '場地'] 
                df_grouped = df_label.groupby(group_cols, dropna=False, as_index=False).agg({
                    '單科標籤': lambda x: '、'.join(sorted(dict.fromkeys([str(i) for i in x if str(i).strip() != ""]), key=natural_sort_key))
                })
                
                df_grouped = df_grouped.rename(columns={'科目數目': '個人考科', '場地': '地點', '單科標籤': '所有考科'})
                
                df_vld = df_grouped[df_grouped['地點'] != ""].copy()
                df_vld['G_W'] = df_vld['年級'].map(grade_weight).fillna(99)
                df_vld['L_W'] = df_vld['地點'].map(loc_weight).fillna(99)
                df_vld['NumSeat'] = pd.to_numeric(df_vld['座號'], errors='coerce').fillna(999)
                df_vld = df_vld.sort_values(by=['G_W', 'L_W', '個人考科', 'NumSeat', '班級'], ascending=[True, True, False, True, True])
                
                f_dfs = []
                for gr in ['一', '二', '三']:
                    for loc, cap in current_targets.items():
                        sub = df_vld[(df_vld['地點'] == loc) & (df_vld['年級'] == gr)].copy()
                        if sub.empty: continue 
                        if len(sub) < cap: 
                            pad = pd.DataFrame([{'地點': loc, '年級': gr}] * (cap - len(sub)))
                            sub = pd.concat([sub, pad], ignore_index=True)
                        sub['序號'] = [f"{i+1:03d}" for i in range(len(sub))]
                        f_dfs.append(sub)
                
                label_cols = ['所有考科', '班級', '年級', '座號', '姓名', '個人考科', '地點', '序號']
                df_rep2 = pd.concat(f_dfs, ignore_index=True) if f_dfs else pd.DataFrame(columns=label_cols)
                for col in label_cols:
                    if col not in df_rep2.columns: df_rep2[col] = ""
                df_rep2 = df_rep2[label_cols]

                # --- 階段三：報表三處理 (考程) ---
                df_exam = df_target.drop_duplicates(subset=['學號', '試卷編號'], keep='first').copy()

                if '試卷編號' in df_exam.columns:
                    clean_papers = df_exam['試卷編號'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'\s+', '', regex=True).str.upper()
                    df_exam = df_exam[~clean_papers.isin(['', 'NAN', 'NONE', '<NA>', 'NULL'])].copy()
                    df_exam['試卷編號'] = clean_papers[~clean_papers.isin(['', 'NAN', 'NONE', '<NA>', 'NULL'])]

                def extract_loc_short(loc):
                    l = str(loc); return '致用' if '致用' in l else '圖書' if '圖書' in l else '電腦' if '電腦' in l else l

                df_exam['比對場地'] = df_exam['場地'].apply(extract_loc_short) 
                df_teacher['比對場地'] = df_teacher['場地'].apply(extract_loc_short)
                df_exam['應到人數'] = df_exam.groupby(['場地', '班級', '科目簡稱'])['學號'].transform('count')
                
                valid_locs_exam = df_exam[df_exam['比對場地'].isin(['致用', '圖書', '電腦'])]
                loc_counts = valid_locs_exam.groupby(['班級', '試卷編號'])['比對場地'].nunique().to_dict()
                
                def determine_group_id(row):
                    c, p, l = str(row['班級']), str(row['試卷編號']), str(row['比對場地'])
                    if loc_counts.get((c, p), 0) <= 1: return p 
                    suffix = ".1" if l == '致用' else ".2" if l == '圖書' else ".3" if l == '電腦' else ""
                    return f"{p}{suffix}"
                    
                df_exam['分組編號'] = df_exam.apply(determine_group_id, axis=1)
                teacher_map = df_teacher.drop_duplicates(subset=['比對年級', '比對場地']).set_index(['比對年級', '比對場地'])['監考教師'].to_dict()
                df_exam['監考教師'] = df_exam.apply(lambda r: teacher_map.get((r['年級'], r['比對場地']), ""), axis=1)

                final_cols = ['班級', '座號', '學號', '姓名', '科目簡稱', '試卷編號', '分組編號', '場地', '授課教師', '監考教師', '時間2', '應到人數']
                for col in final_cols:
                    if col not in df_exam.columns: df_exam[col] = ""

                df_final_exam = df_exam[final_cols].copy()
                
                # ==========================================
                # ⭐ 排序核心修改區：優先排序場地與人數
                # ==========================================
                df_final_exam['G_W'] = df_final_exam['班級'].apply(grade_to_chinese).map(grade_weight).fillna(99)
                df_final_exam['L_W'] = df_final_exam['場地'].map(loc_weight).fillna(99) # 產生場地權重
                
                # 第一順位: 場地(L_W) / 第二順位: 應到人數(預設降冪 False，人多的卷袋排在前面)
                df_final_exam = df_final_exam.sort_values(
                    by=['L_W', '場地', '應到人數', 'G_W', '班級', '科目簡稱', '座號'], 
                    ascending=[True, True, False, True, True, True, True] 
                )

                # 修改 GroupKey 的組合順序，確保與剛剛的排序邏輯相符
                df_final_exam['GroupKey'] = df_final_exam['場地'] + "_" + df_final_exam['班級'] + "_" + df_final_exam['科目簡稱']
                
                grouped = [g for _, g in df_final_exam.groupby('GroupKey', sort=False)]
                final_rows = []
                empty = pd.DataFrame([[np.nan] * len(final_cols)], columns=final_cols)
                
                for i, grp in enumerate(grouped):
                    final_rows.append(grp.drop(columns=['GroupKey', 'G_W', 'L_W'])) # 記得將運算用的輔助欄位丟掉
                    if i < len(grouped) - 1: final_rows.append(empty)
                
                if final_rows:
                    df_rep3_final = pd.concat(final_rows, ignore_index=True).fillna("")
                else:
                    df_rep3_final = pd.DataFrame(columns=final_cols)

                # --- 階段四：報表四處理 (印卷) ---
                df_rep4 = df_target[df_target['試卷編號'] != ""].drop_duplicates(subset=['學號', '試卷編號']).groupby('試卷編號').size().reset_index(name='試卷數量')
                df_rep4['SortKey'] = df_rep4['試卷編號'].apply(natural_sort_key)
                df_rep4 = df_rep4.sort_values(by='SortKey').drop(columns=['SortKey'])

                # --- 欄位重新命名以完美銜接合併列印 ---
                rename_mapping = {'姓名': '學生姓名', '場地': '地點'}
                
                df_target_out = df_target.rename(columns=rename_mapping)
                df_rep2_out = df_rep2.rename(columns=rename_mapping)
                df_rep3_out = df_rep3_final.rename(columns=rename_mapping)
                df_rep4_out = df_rep4.rename(columns=rename_mapping)

                # 【鎖定記憶體】
                st.session_state['results'] = {
                    'venue': to_excel_bytes(df_target_out),
                    'label': to_excel_bytes(df_rep2_out),
                    'schedule': to_excel_bytes(df_rep3_out),
                    'print': to_excel_bytes(df_rep4_out)
                }
                st.balloons()

            except Exception as e:
                st.error("🚨 發生未預期錯誤，請檢查檔案格式是否正確。")
                with st.expander("點此查看詳細工程錯誤碼"):
                    st.code(traceback.format_exc())

# ==========================================
# 5. 下載區
# ==========================================
if st.session_state['results'] is not None:
    st.divider()
    st.success("🎊 運算結果已鎖定，您可以逐一下載所有檔案：")
    
    res = st.session_state['results']
    d_col1, d_col2 = st.columns(2)
    
    with d_col1:
        st.download_button("📄 下載：1.場地分配版", res['venue'], "1_場地分配版.xlsx", "application/vnd.ms-excel", use_container_width=True)
        st.download_button("🖨️ 下載：2.排座標籤", res['label'], "2_報表二_排座標籤.xlsx", "application/vnd.ms-excel", use_container_width=True)
    with d_col2:
        st.download_button("📋 下載：3.考程匯整表", res['schedule'], "3_全校補考考程匯整表.xlsx", "application/vnd.ms-excel", use_container_width=True)
        st.download_button("📝 下載：4.試卷印製表", res['print'], "4_試卷印製數量表.xlsx", "application/vnd.ms-excel", use_container_width=True)
