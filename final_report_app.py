
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="問卷互動分析報告")

@st.cache_data
def load_data(file_paths):
    """Loads and concatenates data from a list of file paths."""
    all_dfs = []
    for path in file_paths:
        try:
            df = pd.read_csv(path)
            all_dfs.append(df)
        except FileNotFoundError:
            st.error(f"錯誤：在應用程式的根目錄中找不到資料檔案 {path}。請確認您已將所有 CSV 檔案和 app 腳本一同上傳至 GitHub。")
            return None
    
    if not all_dfs:
        return None

    merged_df = pd.concat(all_dfs, ignore_index=True, sort=False)
    merged_df.columns = merged_df.columns.str.replace('\n', ' ', regex=False)
    return merged_df

# --- App Header ---
st.title("📊 問卷資料互動分析報告")
st.markdown("請先選擇分析模式，然後再根據提示選擇要查看的資料範圍。")

# --- File Definitions ---
company_files = {
    "第一階段": "STANDARD_8RG8Y_未上市櫃公司治理問卷第一階段_202511050604_690ae8db08878.csv",
    "第二階段": "STANDARD_7RGxP_未上市櫃公司治理問卷第二階段_202511050605_690ae92a9a127.csv",
    "第三階段": "STANDARD_Yb9D2_未上市櫃公司治理問卷第三階段_202511050605_690ae9445a228.csv"
}

investor_files = {
    "第一階段": "STANDARD_NwNYM_未上市櫃公司治理問卷第一階段投資方_202511060133_690bfaccec28e.csv",
    "第二階段": "STANDARD_v2xYO_未上市櫃公司治理問卷第二階段投資方_202511060133_690bfae9b9065.csv",
    "第三階段": "STANDARD_we89e_未上市櫃公司治理問卷第三階段投資方_202511060133_690bfb0524491.csv"
}

# --- Top-level Mode Selection ---
analysis_mode = st.radio(
    "**步驟一：請選擇分析模式**",
    ('分開比較', '合併分析'),
    horizontal=True
)

files_to_load = []
report_title = ""

if analysis_mode == '分開比較':
    data_side = st.radio(
        "**步驟二：請選擇要分析的對象**",
        ('公司方', '投資方'),
        horizontal=True,
        key='data_side_selector'
    )
    
    if data_side == '公司方':
        phases = company_files
    else:
        phases = investor_files

    # Add the new "No Phase" option
    phase_options = list(phases.keys()) + ["不分階段 (全部合併)"]
    selected_phase_name = st.selectbox("**步驟三：請選擇問卷階段**", phase_options, key='phase_selector_separate')

    if selected_phase_name == "不分階段 (全部合併)":
        files_to_load = list(phases.values())
    else:
        files_to_load.append(phases[selected_phase_name])
    
    report_title = f"{data_side} - {selected_phase_name}"

else: # Merged Analysis
    merge_option = st.selectbox("**步驟二：請選擇合併範圍**", (
        "第一階段 (合併)", 
        "第二階段 (合併)", 
        "第三階段 (合併)", 
        "不分階段 (全部合併)"
    ), key='phase_selector_merged')

    if merge_option == "第一階段 (合併)":
        files_to_load = [company_files["第一階段"], investor_files["第一階段"]]
    elif merge_option == "第二階段 (合併)":
        files_to_load = [company_files["第二階段"], investor_files["第二階段"]]
    elif merge_option == "第三階段 (合併)":
        files_to_load = [company_files["第三階段"], investor_files["第三階段"]]
    else: # All
        files_to_load = list(company_files.values()) + list(investor_files.values())
    
    report_title = merge_option

# --- Data Loading & Analysis ---
st.header(f"您正在查看：{report_title}的分析結果")
df = load_data(files_to_load)

if df is not None:
    st.metric(label="總樣本數 (問卷份數)", value=len(df))
    st.markdown("---")

    expand_all = st.checkbox("一鍵展開/收合所有題目", value=False, key="expand_all_toggle")
    st.markdown("---")

    cols_to_exclude = [
        '為了後續支付訪談費，請提供您的電子郵件地址（我們將僅用於聯繫您支付訪談費，並妥善保護您的資料）:', 
        'IP紀錄', '額滿結束註記', '使用者紀錄', '會員時間', 'Hash', '會員編號', '自訂ID', '備註', '填答時間'
    ]
    analysis_cols = [col for col in df.columns if col not in cols_to_exclude]

    for i, col_name in enumerate(analysis_cols):
        with st.expander(f"題目：{col_name}", expanded=expand_all):
            col_data = df[col_name].dropna()

            if col_data.empty:
                st.warning("此欄位無有效資料可供分析。")
                continue

            is_multiselect = False
            if col_data.dtype == 'object':
                non_empty_data = col_data[col_data.astype(str) != '']
                if not non_empty_data.empty and non_empty_data.str.contains('\n').any():
                    is_multiselect = True

            if is_multiselect:
                st.markdown("##### 複選題選項次數分佈")
                exploded_data = col_data.str.split('\n').explode().str.strip()
                exploded_data = exploded_data[exploded_data != '']
                stats_df = exploded_data.value_counts().reset_index()
                stats_df.columns = ['獨立選項', '次數']
                st.dataframe(stats_df)

                st.markdown("##### 垂直長條圖")
                fig = go.Figure(data=[go.Bar(x=stats_df['獨立選項'], y=stats_df['次數'])])
                fig.update_layout(xaxis_tickangle=0, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_multi")

            else:
                is_numeric = pd.api.types.is_numeric_dtype(col_data)
                if not is_numeric:
                    numeric_version = pd.to_numeric(col_data, errors='coerce')
                    if (numeric_version.notna().sum() / len(col_data) > 0.7):
                        is_numeric = True
                        col_data = numeric_version.dropna()

                if is_numeric:
                    st.markdown("##### 數值型資料統計摘要")
                    st.dataframe(col_data.describe().to_frame().T.style.format("{:,.2f}"))
                    st.markdown("##### 盒狀圖")
                    fig = go.Figure(data=[go.Box(y=col_data, name=col_name)])
                    fig.update_layout(xaxis_tickangle=0, template="plotly_white")
                    st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_num")

                else:
                    st.markdown("##### 類別型資料次數分佈")
                    stats_df = col_data.astype(str).value_counts().reset_index()
                    stats_df.columns = ['答案選項', '次數']
                    st.dataframe(stats_df)

                    st.markdown("##### 垂直長條圖")
                    fig = go.Figure(data=[go.Bar(x=stats_df['答案選項'], y=stats_df['次數'])])
                    fig.update_layout(xaxis_tickangle=0, template="plotly_white")
                    st.plotly_chart(fig, use_container_width=True, key=f"plot_{report_title}_{i}_cat")
