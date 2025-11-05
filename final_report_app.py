import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="問卷互動分析報告")

@st.cache_data
def load_phase_data(file_path):
    """Loads and cleans data for a specific phase."""
    try:
        df = pd.read_csv(file_path)
        df.columns = df.columns.str.replace('\n', ' ', regex=False)
        return df
    except FileNotFoundError:
        st.error(f"錯誤：找不到資料檔案 {file_path}")
        return None

st.title("📊 問卷資料互動分析報告 (最終版)")
st.markdown("請從下方的下拉選單中選擇一個問卷階段，以查看該階段所有問題的獨立分析結果。" )

phases = {
    "第一階段": "/Users/liuchenbang/Desktop/工作/STANDARD_8RG8Y_未上市櫃公司治理問卷第一階段_202511050604_690ae8db08878.csv",
    "第二階段": "/Users/liuchenbang/Desktop/工作/STANDARD_7RGxP_未上市櫃公司治理問卷第二階段_202511050605_690ae92a9a127.csv",
    "第三階段": "/Users/liuchenbang/Desktop/工作/STANDARD_Yb9D2_未上市櫃公司治理問卷第三階段_202511050605_690ae9445a228.csv"
}

selected_phase_name = st.selectbox("**請選擇要分析的問卷階段：**", list(phases.keys()))

df = load_phase_data(phases[selected_phase_name])

if df is not None:
    st.header(f"您正在查看：{selected_phase_name}的分析結果")

    cols_to_exclude = [
        '為了後續支付訪談費，請提供您的電子郵件地址（我們將僅用於聯繫您支付訪談費，並妥善保護您的資料）:', 
        'IP紀錄', '額滿結束註記', '使用者紀錄', '會員時間', 'Hash', '會員編號', '自訂ID', '備註', '填答時間'
    ]
    analysis_cols = [col for col in df.columns if col not in cols_to_exclude]

    for i, col_name in enumerate(analysis_cols):
        st.subheader(f"題目：{col_name}")
        
        col_data = df[col_name].dropna()

        if col_data.empty:
            st.warning("此欄位無有效資料可供分析。" )
            st.markdown("---")
            continue

        # Heuristic to detect multi-select questions
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

            st.markdown("##### 垂直長條圖 (Vertical Bar Chart)")
            fig = go.Figure(data=[go.Bar(x=stats_df['獨立選項'], y=stats_df['次數'])])
            fig.update_layout(xaxis_tickangle=0, template="plotly_white")
            st.plotly_chart(fig, use_container_width=True, key=f"multiselect_plot_{selected_phase_name}_{i}")

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
                
                st.markdown("##### 盒狀圖 (Box Plot)")
                fig = go.Figure(data=[go.Box(y=col_data, name=col_name)])
                fig.update_layout(xaxis_tickangle=0, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True, key=f"num_plot_{selected_phase_name}_{i}")

            else:
                st.markdown("##### 類別型資料次數分佈")
                stats_df = col_data.astype(str).value_counts().reset_index()
                stats_df.columns = ['答案選項', '次數']
                st.dataframe(stats_df)

                st.markdown("##### 垂直長條圖 (Vertical Bar Chart)")
                fig = go.Figure(data=[go.Bar(x=stats_df['答案選項'], y=stats_df['次數'])])
                fig.update_layout(xaxis_tickangle=0, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True, key=f"cat_plot_{selected_phase_name}_{i}")

        st.markdown("---")