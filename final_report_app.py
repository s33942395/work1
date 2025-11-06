
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from scipy.stats import ttest_ind, chi2_contingency, f_oneway
import warnings

st.set_page_config(layout="wide", page_title="問卷互動分析報告")

@st.cache_data
def load_and_concat(file_paths):
    """Loads, cleans column names, and concatenates data from a list of file paths."""
    all_dfs = []
    for path in file_paths:
        try:
            df = pd.read_csv(path)
            df.columns = df.columns.str.replace(r'【.*?】', '', regex=True).str.strip()
            df.columns = df.columns.str.replace('\n', ' ', regex=False)
            all_dfs.append(df)
        except FileNotFoundError:
            st.error(f"錯誤：找不到資料檔案 {path}。請確認所有 CSV 檔案都已和 app 腳本一同上傳至 GitHub。")
            return None
    if not all_dfs: return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True, sort=False)

st.title("📊 問卷資料互動分析報告")
st.markdown("請先選擇分析模式，然後再根據提示選擇要查看的資料範圍。")

# --- File Definitions ---
# Using relative paths for deployment
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
analysis_mode = st.radio("**步驟一：請選擇分析模式**", ('逐題瀏覽', '合併分析', '總體統計摘要'), horizontal=True, key="main_mode")

if analysis_mode == '總體統計摘要':
    st.header("總體統計摘要")
    warnings.filterwarnings('ignore')
    
    all_company_df = load_and_concat(list(company_files.values()) + [COMPANY_NEW_MULTIPHASE_FILE])
    all_investor_df = load_and_concat(list(investor_files.values()))
    all_df = pd.concat([all_company_df, all_investor_df], ignore_index=True, sort=False)

    if all_df is not None and not all_df.empty:
        st.metric(label="總樣本數 (問卷份數)", value=len(all_df))
        st.markdown("---")

        with st.expander("1. 數值變數相關性矩陣", expanded=True):
            numeric_cols = all_df.select_dtypes(include=np.number).columns.tolist()
            corr_df = all_df[numeric_cols].corr()
            fig = go.Figure(data=go.Heatmap(z=corr_df.values, x=corr_df.columns, y=corr_df.columns, colorscale='Blues'))
            fig.update_layout(title='數值變數之間的相關性')
            st.plotly_chart(fig, use_container_width=True, key="corr_matrix")

        with st.expander("2. 公司方 vs. 投資方 差異檢定 (T-test)", expanded=True):
            num_col_to_test = '請問公司的實收資本額：'
            if num_col_to_test in all_company_df.columns and num_col_to_test in all_investor_df.columns:
                company_vals = pd.to_numeric(all_company_df[num_col_to_test], errors='coerce').dropna()
                investor_vals = pd.to_numeric(all_investor_df[num_col_to_test], errors='coerce').dropna()
                if len(company_vals) > 5 and len(investor_vals) > 5:
                    stat, p_value = ttest_ind(company_vals, investor_vals, equal_var=False)
                    st.markdown(f"- **檢定變數**: `{num_col_to_test.strip('：')}`")
                    st.markdown(f"- **檢定統計量 (T-statistic)**: `{stat:.4f}`")
                    st.markdown(f"- **p-value**: `{p_value:.4f}`")
                    st.markdown(f"- **結論**: {'**存在**顯著差異' if p_value < 0.05 else '**未發現**顯著差異'}")
            else: st.warning("公司方或投資方資料中缺少「實收資本額」欄位，無法執行 T-test。")

        with st.expander("3. 不同發展階段公司差異檢定 (ANOVA)", expanded=True):
            df_new_multi = load_and_concat([COMPANY_NEW_MULTIPHASE_FILE])
            if df_new_multi is not None and not df_new_multi.empty:
                p1_data = df_new_multi[df_new_multi[PHASE_COLUMN_NAME].str.contains("第一階段", na=False)]
                p2_data = df_new_multi[df_new_multi[PHASE_COLUMN_NAME].str.contains("第二階段", na=False)]
                anova_col = '請問公司的實收資本額：'
                if anova_col in p1_data.columns and anova_col in p2_data.columns:
                    group1 = pd.to_numeric(p1_data[anova_col], errors='coerce').dropna()
                    group2 = pd.to_numeric(p2_data[anova_col], errors='coerce').dropna()
                    if len(group1) > 1 and len(group2) > 1:
                        f_stat, p_value = f_oneway(group1, group2)
                        st.markdown(f"- **檢定變數**: `{anova_col.strip('：')}`")
                        st.markdown(f"- **檢定統計量 (F-statistic)**: `{f_stat:.4f}`")
                        st.markdown(f"- **p-value**: `{p_value:.4f}`")
                        st.markdown(f"- **結論**: 比較第一和第二階段，其平均實收資本額{'**存在**顯著差異' if p_value < 0.05 else '**未發現**顯著差異'}")
            else: st.warning("新公司方檔案中無足夠的階段資料可進行 ANOVA 檢定。")
    else: st.warning("無足夠資料可進行總體統計分析。")

else: # Detailed question-by-question browser
    df_to_analyze = pd.DataFrame()
    report_title = ""
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
            df_new_multi = load_and_concat([COMPANY_NEW_MULTIPHASE_FILE])
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
    
    elif analysis_mode == '合併分析':
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
        if phase_filter:
            df_new_multi = load_and_concat([COMPANY_NEW_MULTIPHASE_FILE])
            if df_new_multi is not None and not df_new_multi.empty:
                df_filtered = df_new_multi[df_new_multi[PHASE_COLUMN_NAME].str.contains(phase_filter, na=False)]
                df_list.append(df_filtered)
        if df_list: df_to_analyze = pd.concat(df_list, ignore_index=True, sort=False)

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
            with st.expander(f"題目：{col_name}", expanded=expand_all):
                col_data = df_to_analyze[col_name].dropna()
                if col_data.empty: continue
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
    else: st.warning("在此選擇下沒有載入任何資料，請檢查您的選擇和檔案。")
