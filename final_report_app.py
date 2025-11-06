
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="問卷互動分析報告")

@st.cache_data
def load_and_concat(file_paths):
    """Loads, cleans column names, and concatenates data from a list of file paths."""
    all_dfs = []
    for path in file_paths:
        try:
            df = pd.read_csv(path)
            # Normalize column names by removing prefixes like 【...】 and stripping whitespace
            df.columns = df.columns.str.replace(r'【.*?】', '', regex=True).str.strip()
            df.columns = df.columns.str.replace('\n', ' ', regex=False)
            all_dfs.append(df)
        except FileNotFoundError:
            st.error(f"錯誤：找不到資料檔案 {path}。請確認所有 CSV 檔案都已和 app 腳本一同上傳至 GitHub。")
            return None
    
    if not all_dfs:
        return pd.DataFrame()

    merged_df = pd.concat(all_dfs, ignore_index=True, sort=False)
    return merged_df

# --- App Header and File Definitions ---
st.title("📊 問卷資料互動分析報告")
st.markdown("請先選擇分析模式，然後再根據提示選擇要查看的資料範圍。")

# --- Use RELATIVE paths for deployment ---
COMPANY_P1_FILE = "STANDARD_8RG8Y_未上市櫃公司治理問卷第一階段_202511050604_690ae8db08878.csv"
COMPANY_P2_FILE = "STANDARD_7RGxP_未上市櫃公司治理問卷第二階段_202511050605_690ae92a9a127.csv"
COMPANY_P3_FILE = "STANDARD_Yb9D2_未上市櫃公司治理問卷第三階段_202511050605_690ae9445a228.csv"
INVESTOR_P1_FILE = "STANDARD_NwNYM_未上市櫃公司治理問卷第一階段投資方_202511060133_690bfaccec28e.csv"
INVESTOR_P2_FILE = "STANDARD_v2xYO_未上市櫃公司治理問卷第二階段投資方_202511060133_690bfae9b9065.csv"
INVESTOR_P3_FILE = "STANDARD_we89e_未上市櫃公司治理問卷第三階段投資方_202511060133_690bfb0524491.csv"
COMPANY_NEW_MULTIPHASE_FILE = "STANDARD_v2xkX_未上市櫃公司治理問卷_202511060532_690c3305c62b5.csv"
PHASE_COLUMN_NAME = "請問公司目前主要處於哪個發展階段？："

company_files = {"第一階段": COMPANY_P1_FILE, "第二階段": COMPANY_P2_FILE, "第三階段": COMPANY_P3_FILE}
investor_files = {"第一階段": INVESTOR_P1_FILE, "第二階段": INVESTOR_P2_FILE, "第三階段": INVESTOR_P3_FILE}

# --- UI Logic ---
analysis_mode = st.radio("**步驟一：請選擇分析模式**", ('分開比較', '合併分析'), horizontal=True)

report_title = ""
df_to_analyze = pd.DataFrame()

# We need to load the new multi-phase file to filter it, so load it once here.
try:
    df_new_multi = load_and_concat([COMPANY_NEW_MULTIPHASE_FILE])
except Exception as e:
    # This will be caught by the load_and_concat function, but as a fallback:
    st.error(f"無法讀取新的公司方問卷檔案: {COMPANY_NEW_MULTIPHASE_FILE}。請確認此檔案已上傳。")
    df_new_multi = pd.DataFrame()

if analysis_mode == '分開比較':
    data_side = st.radio("**步驟二：請選擇要分析的對象**", ('公司方', '投資方'), horizontal=True, key='data_side_selector')
    phase_options = list(company_files.keys()) + ["不分階段 (全部合併)"]
    selected_phase = st.selectbox("**步驟三：請選擇問卷階段**", phase_options, key='phase_selector_separate')
    report_title = f"{data_side} - {selected_phase}"

    df_list = []
    if data_side == '公司方':
        files_to_load = []
        if selected_phase in company_files:
            files_to_load.append(company_files[selected_phase])
        elif selected_phase == "不分階段 (全部合併)":
            files_to_load.extend(list(company_files.values()))
        
        if files_to_load: df_list.append(load_and_concat(files_to_load))

        if df_new_multi is not None and not df_new_multi.empty:
            if selected_phase in company_files:
                df_filtered = df_new_multi[df_new_multi[PHASE_COLUMN_NAME].str.contains(selected_phase, na=False)]
                df_list.append(df_filtered)
            elif selected_phase == "不分階段 (全部合併)":
                df_list.append(df_new_multi)
    else: # Investor side
        files_to_load = []
        if selected_phase in investor_files:
            files_to_load.append(investor_files[selected_phase])
        else: 
            files_to_load = list(investor_files.values())
        if files_to_load: df_list.append(load_and_concat(files_to_load))
    
    if df_list: df_to_analyze = pd.concat(df_list, ignore_index=True, sort=False)

else: # Merged Analysis
    merge_option = st.selectbox("**步驟二：請選擇合併範圍**", ("第一階段 (合併)", "第二階段 (合併)", "第三階段 (合併)", "不分階段 (全部合併)"), key='phase_selector_merged')
    report_title = merge_option
    files_to_load = []
    phase_filter = None

    if merge_option == "第一階段 (合併)":
        files_to_load = [COMPANY_P1_FILE, INVESTOR_P1_FILE]
        phase_filter = "第一階段"
    elif merge_option == "第二階段 (合併)":
        files_to_load = [COMPANY_P2_FILE, INVESTOR_P2_FILE]
        phase_filter = "第二階段"
    elif merge_option == "第三階段 (合併)":
        files_to_load = [COMPANY_P3_FILE, INVESTOR_P3_FILE]
    else: # All
        files_to_load = list(company_files.values()) + list(investor_files.values()) + [COMPANY_NEW_MULTIPHASE_FILE]

    df_base = load_and_concat(files_to_load)
    df_list = [df_base]
    if phase_filter and df_new_multi is not None and not df_new_multi.empty:
        df_filtered = df_new_multi[df_new_multi[PHASE_COLUMN_NAME].str.contains(phase_filter, na=False)]
        df_list.append(df_filtered)
    
    if df_list: df_to_analyze = pd.concat(df_list, ignore_index=True, sort=False)

# --- Display Analysis ---
st.header(f"您正在查看：{report_title}的分析結果")
if df_to_analyze is not None and not df_to_analyze.empty:
    st.metric(label="總樣本數 (問卷份數)", value=len(df_to_analyze))
    st.markdown("---")
    expand_all = st.checkbox("一鍵展開/收合所有題目", value=False, key="expand_all_toggle")
    st.markdown("---")

    cols_to_exclude = ['為了後續支付訪談費，請提供您的電子郵件地址（我們將僅用於聯繫您支付訪談費，並妥善保護您的資料）:', 'IP紀錄', '額滿結束註記', '使用者紀錄', '會員時間', 'Hash', '會員編號', '自訂ID', '備註', '填答時間', PHASE_COLUMN_NAME]
    analysis_cols = [col for col in df_to_analyze.columns if col not in cols_to_exclude and col in df_to_analyze.columns]
    analysis_cols = list(pd.Series(analysis_cols)) # Get unique columns while preserving order

    for i, col_name in enumerate(analysis_cols):
        with st.expander(f"題目：{col_name}", expanded=expand_all):
            col_data = df_to_analyze[col_name].dropna()
            if col_data.empty: st.warning("此欄位無有效資料可供分析。"); continue
            is_multiselect = False
            if col_data.dtype == 'object':
                non_empty_data = col_data[col_data.astype(str) != '']
                if not non_empty_data.empty and non_empty_data.str.contains('\n').any(): is_multiselect = True
            if is_multiselect:
                st.markdown("##### 複選題選項次數分佈"); exploded_data = col_data.str.split('\n').explode().str.strip(); exploded_data = exploded_data[exploded_data != '']; stats_df = exploded_data.value_counts().reset_index(); stats_df.columns = ['獨立選項', '次數']; st.dataframe(stats_df)
                st.markdown("##### 垂直長條圖"); fig = go.Figure(data=[go.Bar(x=stats_df['獨立選項'], y=stats_df['次數'])]); fig.update_layout(xaxis_tickangle=0, template="plotly_white"); st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_multi")
            else:
                is_numeric = pd.api.types.is_numeric_dtype(col_data)
                if not is_numeric:
                    numeric_version = pd.to_numeric(col_data, errors='coerce');
                    if (numeric_version.notna().sum() / len(col_data) > 0.7): is_numeric = True; col_data = numeric_version.dropna()
                if is_numeric:
                    st.markdown("##### 數值型資料統計摘要"); st.dataframe(col_data.describe().to_frame().T.style.format("{:,.2f}")); st.markdown("##### 盒狀圖"); fig = go.Figure(data=[go.Box(y=col_data, name=col_name)]); fig.update_layout(xaxis_tickangle=0, template="plotly_white"); st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_num")
                else:
                    st.markdown("##### 類別型資料次數分佈"); stats_df = col_data.astype(str).value_counts().reset_index(); stats_df.columns = ['答案選項', '次數']; st.dataframe(stats_df)
                    st.markdown("##### 垂直長條圖"); fig = go.Figure(data=[go.Bar(x=stats_df['答案選項'], y=stats_df['次數'])]); fig.update_layout(xaxis_tickangle=0, template="plotly_white"); st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_cat")
else:
    st.warning("沒有載入任何資料，請檢查檔案路徑和選擇的選項。")
