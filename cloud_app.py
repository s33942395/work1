import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import warnings
import os
import re

st.set_page_config(layout="wide", page_title="問卷互動分析報告")

@st.cache_data
def load_and_concat(file_paths):
    all_dfs = []
    for path in file_paths:
        try:
            df = pd.read_csv(path)
            df.columns = df.columns.str.replace(r'【.*?】', '', regex=True).str.strip()
            df.columns = df.columns.str.replace('\n', ' ', regex=False)

            # 若 CSV 已有階段欄位，將像 "第一階段：..."、"第一階段／..." 等值正規化為 "第一階段"
            if PHASE_COLUMN_NAME in df.columns:
                extracted = df[PHASE_COLUMN_NAME].astype(str).str.extract(r'(第一階段|第二階段|第三階段)', expand=False)
                # 若能抽出標準階段名稱，使用它；否則保留原值（避免破壞非預期格式）
                df[PHASE_COLUMN_NAME] = extracted.where(extracted.notna(), df[PHASE_COLUMN_NAME])
            else:
                # 如果 CSV 本身沒有階段欄位，嘗試從檔名推斷（第一階段 / 第二階段 / 第三階段）
                m = re.search(r'(第一階段|第二階段|第三階段)', os.path.basename(path))
                if m:
                    df[PHASE_COLUMN_NAME] = m.group(1)

            # 加入來源檔名以便追蹤來源
            df['_source_file'] = os.path.basename(path)

            all_dfs.append(df)
        except FileNotFoundError:
            st.error(f"錯誤：找不到資料檔案 {path}。請確認所有 CSV 檔案都已和 app 腳本一同上傳至 GitHub。")
            return None
    if not all_dfs: return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True, sort=False)

st.title("📊 問卷資料互動分析報告 (雲端版)")
st.markdown("請先選擇分析模式，然後再根據提示選擇要查看的資料範圍。")

# --- File Definitions (Relative Paths for Cloud) ---
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
analysis_mode = st.radio("**步驟一：請選擇分析模式**", ('逐題瀏覽', '合併分析'), horizontal=True, key="main_mode")

df_to_analyze = pd.DataFrame()
report_title = ""

try:
    df_new_multi = load_and_concat([COMPANY_NEW_MULTIPHASE_FILE])
except Exception:
    df_new_multi = pd.DataFrame()

if analysis_mode == '逐題瀏覽':
    data_side = st.radio("**步驟二：請選擇要分析的對象**", ('公司方', '投資方'), horizontal=True, key='data_side_selector')
    phase_options = list(company_files.keys()) + ["不分階段 (全部合併)"]
    selected_phase = st.selectbox("**步驟三：請選擇問卷階段**", phase_options, key='phase_selector_separate')
    report_title = f"{data_side} - {selected_phase}"
    df_list = []
    if data_side == '公司方':
        files_to_load = []
        if selected_phase in company_files: files_to_load.append(company_files[selected_phase])
        elif selected_phase == "不分階段 (全部合併)": files_to_load.extend(list(company_files.values()))
        if files_to_load: df_list.append(load_and_concat(files_to_load))
        if df_new_multi is not None and not df_new_multi.empty:
            if selected_phase in company_files: 
                df_filtered = df_new_multi[df_new_multi[PHASE_COLUMN_NAME].str.contains(selected_phase, na=False)]
                df_list.append(df_filtered)
            elif selected_phase == "不分階段 (全部合併)": df_list.append(df_new_multi)
    else: # Investor side
        files_to_load = []
        if selected_phase in investor_files: files_to_load.append(investor_files[selected_phase])
        else: files_to_load = list(investor_files.values())
        if files_to_load: df_list.append(load_and_concat(files_to_load))
    if df_list: df_to_analyze = pd.concat(df_list, ignore_index=True, sort=False)

else: # Merged Analysis
    merge_option = st.selectbox("**步驟二：請選擇合併範圍**", ("第一階段 (合併)", "第二階段 (合併)", "第三階段 (合併)", "不分階段 (全部合併)"), key='phase_selector_merged')
    report_title = merge_option
    files_to_load = []
    phase_filter = None
    if merge_option == "第一階段 (合併)": files_to_load, phase_filter = [COMPANY_P1_FILE, INVESTOR_P1_FILE], "第一階段"
    elif merge_option == "第二階段 (合併)": files_to_load, phase_filter = [COMPANY_P2_FILE, INVESTOR_P2_FILE], "第二階段"
    elif merge_option == "第三階段 (合併)": files_to_load = [COMPANY_P3_FILE, INVESTOR_P3_FILE]
    else: files_to_load = list(company_files.values()) + list(investor_files.values()) + [COMPANY_NEW_MULTIPHASE_FILE]
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
    analysis_cols = list(pd.Series(analysis_cols))
    for i, col_name in enumerate(analysis_cols):
        col_data = df_to_analyze[col_name].dropna()
        if not col_data.empty:
            with st.expander(f"題目：{col_name}", expanded=expand_all):
                is_multiselect = False
                if col_data.dtype == 'object':
                    non_empty_data = col_data[col_data.astype(str) != '']
                    if not non_empty_data.empty and non_empty_data.str.contains('\n').any(): is_multiselect = True
                if is_multiselect:
                    st.markdown("##### 複選題選項次數分佈")
                    exploded = col_data.astype(str).str.split('\n').explode().str.strip()
                    exploded = exploded[exploded != '']
                    total_counts = exploded.value_counts().reset_index()
                    total_counts.columns = ['獨立選項', '次數']
                    st.dataframe(total_counts)

                    # 若資料含階段欄位且有多個階段，則按階段分色（堆疊）
                    if PHASE_COLUMN_NAME in df_to_analyze.columns and df_to_analyze[PHASE_COLUMN_NAME].notna().any() and df_to_analyze[PHASE_COLUMN_NAME].nunique() > 1:
                        st.markdown("##### 各階段分佈（不同顏色代表不同階段，採堆疊顯示）")
                        exploded_df = exploded.to_frame(name='option')
                        exploded_df['phase'] = df_to_analyze.loc[exploded_df.index, PHASE_COLUMN_NAME].fillna('未標註階段')
                        pivot = exploded_df.groupby(['option', 'phase']).size().unstack(fill_value=0)
                        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                        fig = go.Figure()
                        for j, phase in enumerate(pivot.columns):
                            fig.add_trace(go.Bar(x=pivot.index, y=pivot[phase], name=str(phase), marker_color=colors[j % len(colors)]))
                        fig.update_layout(barmode='stack', xaxis_tickangle=0, template="plotly_white")
                        st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_multi")
                    else:
                        st.markdown("##### 垂直長條圖")
                        fig = go.Figure(data=[go.Bar(x=total_counts['獨立選項'], y=total_counts['次數'])])
                        fig.update_layout(xaxis_tickangle=0, template="plotly_white")
                        st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_multi")
                else:
                    is_numeric = pd.api.types.is_numeric_dtype(col_data)
                    if not is_numeric:
                        numeric_version = pd.to_numeric(col_data, errors='coerce');
                        if (numeric_version.notna().sum() / len(col_data) > 0.7): is_numeric = True; col_data = numeric_version.dropna()
                    if is_numeric:
                        st.markdown("##### 數值型資料統計摘要"); st.dataframe(col_data.describe().to_frame().T.style.format("{:,.2f}")); st.markdown("##### 盒狀圖"); fig = go.Figure(data=[go.Box(y=col_data, name=col_name)]); fig.update_layout(xaxis_tickangle=0, template="plotly_white"); st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_num")
                    else:
                        st.markdown("##### 類別型資料次數分佈")
                        s = col_data.astype(str)
                        total = s.value_counts().reset_index()
                        total.columns = ['答案選項', '次數']
                        st.dataframe(total)

                        # 若資料含階段欄位且有多個階段，則按階段分色（堆疊）
                        if PHASE_COLUMN_NAME in df_to_analyze.columns and df_to_analyze[PHASE_COLUMN_NAME].notna().any() and df_to_analyze[PHASE_COLUMN_NAME].nunique() > 1:
                            st.markdown("##### 各階段分佈（不同顏色代表不同階段，採堆疊顯示）")
                            df_pair = s.to_frame(name='ans')
                            df_pair['phase'] = df_to_analyze.loc[df_pair.index, PHASE_COLUMN_NAME].fillna('未標註階段')
                            pivot = df_pair.groupby(['ans', 'phase']).size().unstack(fill_value=0)
                            colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                            fig = go.Figure()
                            for j, phase in enumerate(pivot.columns):
                                fig.add_trace(go.Bar(x=pivot.index, y=pivot[phase], name=str(phase), marker_color=colors[j % len(colors)]))
                            fig.update_layout(barmode='stack', xaxis_tickangle=0, template="plotly_white")
                            st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_cat")
                        else:
                            st.markdown("##### 垂直長條圖")
                            fig = go.Figure(data=[go.Bar(x=total['答案選項'], y=total['次數'])])
                            fig.update_layout(xaxis_tickangle=0, template="plotly_white")
                            st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_cat")
else: st.warning("在此選擇下沒有載入任何資料，請檢查您的選擇和檔案。")