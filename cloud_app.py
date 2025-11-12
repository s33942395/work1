import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import warnings
import os
import re
from scipy.stats import chi2_contingency, kruskal, mannwhitneyu, fisher_exact
from datetime import datetime
import io
from difflib import SequenceMatcher

warnings.filterwarnings('ignore')

# --- 智慧排序函式 ---
def smart_sort_categories(categories):
    """
    智慧排序類別資料，處理：
    1. 百分比範圍 (如 10-20%, 20-30%)
    2. 數值範圍 (如 1-5年, 5-10年)
    3. 金額範圍 (如 100-500萬, 500-1000萬)
    4. 階段 (第一階段, 第二階段, 第三階段)
    5. 一般文字 (按原順序或字母排序)
    """
    if len(categories) == 0:
        return []
    
    categories_list = list(categories)
    
    # 定義排序鍵函式
    def sort_key(item):
        item_str = str(item).strip()
        
        # 1. 處理百分比範圍 (如 10-20%, 20%-30%)
        percent_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*[%％]', item_str)
        if percent_match:
            return (0, float(percent_match.group(1)))
        
        # 單一百分比 (如 30%)
        single_percent = re.match(r'(\d+\.?\d*)\s*[%％]', item_str)
        if single_percent:
            return (0, float(single_percent.group(1)))
        
        # 2. 處理年份範圍 (如 1-5年, 5-10年)
        year_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*年', item_str)
        if year_match:
            return (1, float(year_match.group(1)))
        
        # 3. 處理金額範圍 (如 100-500萬, 1000-5000萬)
        money_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*[萬億]', item_str)
        if money_match:
            return (2, float(money_match.group(1)))
        
        # 4. 處理月份範圍 (如 1-3個月, 3-6個月)
        month_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*個?月', item_str)
        if month_match:
            return (3, float(month_match.group(1)))
        
        # 5. 處理人數範圍 (如 1-10人, 10-50人)
        people_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*人', item_str)
        if people_match:
            return (4, float(people_match.group(1)))
        
        # 6. 處理次數 (如 每月1次, 每季1次, 每年1次)
        freq_order = {'每週': 1, '每月': 2, '每季': 3, '每半年': 4, '每年': 5, '不定期': 6, '無': 7}
        for key, value in freq_order.items():
            if key in item_str:
                return (5, value)
        
        # 7. 處理階段 (第一階段, 第二階段, 第三階段)
        stage_match = re.search(r'[第]?([一二三四五1234])[階段期]', item_str)
        if stage_match:
            stage_num = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '1': 1, '2': 2, '3': 3, '4': 4}.get(stage_match.group(1), 0)
            return (6, stage_num)
        
        # 8. 處理程度 (完全沒有, 部分有, 完全有)
        degree_order = {
            '完全沒有': 1, '沒有': 1, '無': 1,
            '極少': 2, '很少': 2,
            '部分': 3, '部分有': 3, '部份': 3,
            '大部分': 4, '大部分有': 4,
            '完全': 5, '完全有': 5, '有': 5, '是': 5
        }
        for key, value in degree_order.items():
            if key in item_str:
                return (7, value)
        
        # 9. 處理比較級 (低於, 符合, 高於)
        compare_order = {'低於': 1, '低': 1, '符合': 2, '相當': 2, '高於': 3, '高': 3, '超過': 3}
        for key, value in compare_order.items():
            if key in item_str:
                return (8, value)
        
        # 10. 處理純數字開頭
        num_match = re.match(r'^(\d+\.?\d*)', item_str)
        if num_match:
            return (9, float(num_match.group(1)))
        
        # 11. 特殊處理：「以上」應該排在最後
        if '以上' in item_str or '或以上' in item_str or '以上' in item_str:
            # 提取數字
            num_in_above = re.search(r'(\d+\.?\d*)', item_str)
            if num_in_above:
                return (10, float(num_in_above.group(1)))
        
        # 12. 預設：按字典順序
        return (99, item_str)
    
    # 執行排序
    try:
        sorted_categories = sorted(categories_list, key=sort_key)
        return sorted_categories
    except:
        # 如果排序失敗，返回原順序
        return categories_list

# --- 統計函式定義 ---
def format_p_value(p):
    """顯著性標記"""
    if p < 0.001:
        return f"**p={p:.4f} ⭐⭐⭐ (極顯著)**"
    elif p < 0.01:
        return f"**p={p:.4f} ⭐⭐ (非常顯著)**"
    elif p < 0.05:
        return f"**p={p:.4f} ⭐ (顯著)**"
    else:
        return f"p={p:.4f} (不顯著)"

def interpret_effect_size(cramers_v=None, cohens_d=None):
    """效果量解釋"""
    if cramers_v is not None:
        if cramers_v < 0.1:
            return "效果量極小 (negligible)"
        elif cramers_v < 0.3:
            return "效果量小 (small)"
        elif cramers_v < 0.5:
            return "效果量中等 (medium)"
        else:
            return "效果量大 (large)"
    elif cohens_d is not None:
        if abs(cohens_d) < 0.2:
            return "效果量極小 (negligible)"
        elif abs(cohens_d) < 0.5:
            return "效果量小 (small)"
        elif abs(cohens_d) < 0.8:
            return "效果量中等 (medium)"
        else:
            return "效果量大 (large)"
    return ""

def generate_academic_conclusion(test_type, p_value, effect_size=None, groups_info=None, question_name=""):
    """生成學術風格結論"""
    conclusion = f"\n**📊 學術分析結論 - {question_name}**\n\n"
    
    if test_type == "chi_square":
        conclusion += f"**研究方法：** 採用卡方檢定 (Chi-square test) 檢驗類別變項間的關聯性。\n\n"
        if p_value < 0.05:
            conclusion += f"**研究發現：** 統計結果顯示組間差異達到顯著水準 ({format_p_value(p_value)})，"
            if effect_size:
                conclusion += f"Cramér's V = {effect_size:.3f} ({interpret_effect_size(cramers_v=effect_size)})。"
            conclusion += f"\n\n**實務意涵：** 公司方與投資方在此議題上存在顯著差異，建議進一步探討差異來源。"
        else:
            conclusion += f"**研究發現：** 統計結果顯示組間差異未達顯著水準 ({format_p_value(p_value)})。\n\n"
            conclusion += f"**實務意涵：** 公司方與投資方在此議題上看法趨於一致。"
    
    elif test_type == "mann_whitney":
        conclusion += f"**研究方法：** 採用 Mann-Whitney U 檢定（無母數檢定）比較兩組中位數差異。\n\n"
        if p_value < 0.05:
            conclusion += f"**研究發現：** 統計結果顯示組間差異達到顯著水準 ({format_p_value(p_value)})，"
            if effect_size:
                conclusion += f"Cohen's d = {effect_size:.3f} ({interpret_effect_size(cohens_d=effect_size)})。"
            conclusion += f"\n\n"
            if groups_info:
                conclusion += f"**描述統計：**\n"
                for group, stats in groups_info.items():
                    conclusion += f"- {group}: 中位數={stats['median']:.2f}, 平均數={stats['mean']:.2f}, 標準差={stats['std']:.2f} (n={stats['n']})\n"
            conclusion += f"\n**實務意涵：** 兩組在此議題上存在顯著差異，建議針對差異來源進行深入探討。"
        else:
            conclusion += f"**研究發現：** 統計結果顯示組間差異未達顯著水準 ({format_p_value(p_value)})。\n\n"
            conclusion += f"**實務意涵：** 兩組在此議題上的看法相對一致。"
    
    elif test_type == "kruskal":
        conclusion += f"**研究方法：** 採用 Kruskal-Wallis H 檢定（無母數檢定）比較多組中位數差異。\n\n"
        if p_value < 0.05:
            conclusion += f"**研究發現：** 統計結果顯示組間差異達到顯著水準 ({format_p_value(p_value)})。\n\n"
            if groups_info:
                conclusion += f"**描述統計：**\n"
                for group, stats in groups_info.items():
                    conclusion += f"- {group}: 中位數={stats['median']:.2f}, 平均數={stats['mean']:.2f}, 標準差={stats['std']:.2f} (n={stats['n']})\n"
            conclusion += f"\n**實務意涵：** 不同群體在此議題上的認知或態度存在顯著差異，建議針對差異較大的群體設計差異化策略。"
        else:
            conclusion += f"**研究發現：** 統計結果顯示組間差異未達顯著水準 ({format_p_value(p_value)})。\n\n"
            conclusion += f"**實務意涵：** 各群體在此議題上的看法相對一致。"
    
    elif test_type == "fisher":
        conclusion += f"**研究方法：** 採用 Fisher's Exact Test（適用於小樣本）檢驗類別變項關聯性。\n\n"
        if p_value < 0.05:
            conclusion += f"**研究發現：** 統計結果顯示組間差異達到顯著水準 ({format_p_value(p_value)})。\n\n"
            conclusion += f"**實務意涵：** 儘管樣本數較少，但仍觀察到顯著差異，建議擴大樣本進一步驗證。"
        else:
            conclusion += f"**研究發現：** 統計結果顯示組間差異未達顯著水準 ({format_p_value(p_value)})。"
    
    elif test_type == "multiselect_chi":
        conclusion += f"**研究方法：** 採用 Presence/Absence 卡方檢定分析複選題各選項的組間差異。\n\n"
        conclusion += f"**研究發現：** 請參考下方各選項的統計檢定結果。顯著選項代表該面向在不同群體間有明顯差異。\n\n"
        conclusion += f"**實務意涵：** 建議針對顯著差異的選項，深入探討其背後原因，並考慮調整相應政策或溝通策略。"
    
    return conclusion

def _cramers_v_from_table(table):
    try:
        chi2, p, dof, exp = chi2_contingency(table)
        if np.nanmin(exp) <= 1:
            return None, None, exp
        n = table.values.sum()
        return np.sqrt(chi2 / (n * (min(table.shape) - 1))), p, exp
    except Exception:
        return None, None, None

def compute_and_display_categorical_stats(df, series):
    if PHASE_COLUMN_NAME in df.columns and df[PHASE_COLUMN_NAME].notna().any() and df[PHASE_COLUMN_NAME].nunique() > 1:
        phases = df[PHASE_COLUMN_NAME].fillna('未標註階段')
        table = pd.crosstab(series.astype(str), phases)
        st.markdown("**跨階段統計（類別）**")
        st.dataframe(table)
        cramers, p, exp = _cramers_v_from_table(table)
        if exp is None:
            st.write("無法計算卡方檢定（發生錯誤）。")
        else:
            if np.nanmin(exp) <= 1:
                st.write("卡方檢定未執行：某些 cell 的期望次數 ≤ 1。建議合併類別或改用其他檢定方法。")
                st.write("期望次數矩陣：")
                st.dataframe(pd.DataFrame(exp, index=table.index, columns=table.columns))
            else:
                if p is not None:
                    st.write(f"卡方檢定 {format_p_value(p)}；Cramer's V = {cramers:.3f} ({interpret_effect_size(cramers_v=cramers)})")
                else:
                    st.write("無法計算卡方檢定結果。")
    else:
        st.write("未包含多個階段，未進行跨階段類別檢定。")

def compute_and_display_numeric_stats(df, series):
    if PHASE_COLUMN_NAME in df.columns and df[PHASE_COLUMN_NAME].notna().any() and df[PHASE_COLUMN_NAME].nunique() > 1:
        phases = df[PHASE_COLUMN_NAME].fillna('未標註階段')
        groups = []
        labels = []
        for ph in phases.unique():
            grp = pd.to_numeric(series[phases == ph].dropna(), errors='coerce').dropna().astype(float)
            if len(grp) > 0:
                groups.append(grp)
                labels.append(ph)
        st.markdown("**跨階段統計（數值）**")
        summaries = {lab: f"n={len(g)}, mean={g.mean():.3f}, median={g.median():.3f}, std={g.std(ddof=0):.3f}" for lab, g in zip(labels, groups)}
        st.write(summaries)
        if len(groups) > 1:
            try:
                all_vals = np.concatenate([g.values for g in groups]) if groups else np.array([])
                if all_vals.size > 0 and np.all(all_vals == all_vals[0]):
                    st.write("所有組別的數值完全相同，Kruskal-Wallis 檢定不適用。")
                else:
                    stat, p = kruskal(*groups)
                    st.write(f"Kruskal-Wallis stat={stat:.3f}, {format_p_value(p)}")
            except ValueError as e:
                st.write("Kruskal-Wallis 檢定錯誤：", e)
            except Exception as e:
                st.write("執行 Kruskal-Wallis 檢定時發生錯誤：", e)
        else:
            st.write("每個階段樣本不足，無法進行 Kruskal-Wallis 檢定。")
    else:
        st.write("未包含多個階段，未進行跨階段數值檢定。")

def compute_and_display_multiselect_option_tests(df, original_series, option_list):
    if PHASE_COLUMN_NAME in df.columns and df[PHASE_COLUMN_NAME].notna().any() and df[PHASE_COLUMN_NAME].nunique() > 1:
        st.markdown("**複選題選項跨階段統計（Presence/Absence 卡方）**")
        phases = df[PHASE_COLUMN_NAME].fillna('未標註階段')
        for opt in option_list:
            pres = original_series.astype(str).fillna('').apply(lambda s: opt in [x.strip() for x in s.split('\n') if x.strip()!=''])
            table = pd.crosstab(pres, phases)
            if table.size == 0 or table.values.sum() == 0 or table.shape[0] < 2:
                st.write(f"選項 '{opt}'：樣本或分類不足，無法進行卡方檢定。")
                continue
            try:
                chi2, p, dof, exp = chi2_contingency(table)
                if np.nanmin(exp) <= 1:
                    st.write(f"選項 '{opt}'：期望次數過小 (≤1)，跳過檢定。")
                else:
                    n = table.values.sum()
                    cramers = np.sqrt(chi2 / (n * (min(table.shape) - 1))) if n and min(table.shape) > 1 else None
                    st.write(f"選項 '{opt}'：{format_p_value(p)}" + (f"；Cramer's V={cramers:.3f} ({interpret_effect_size(cramers_v=cramers)})" if cramers is not None else ""))
            except Exception as e:
                st.write(f"選項 '{opt}' 無法計算卡方檢定：{e}")
    else:
        st.write("未包含多個階段，未進行複選題跨階段檢定。")

def perform_comprehensive_statistical_analysis(df, col_data, col_name, is_numeric=False, is_multiselect=False):
    """
    綜合統計分析：分析公司方 vs 投資方、不同階段之間的差異
    """
    st.markdown("---")
    st.markdown("### 📈 統計分析報告")
    
    has_respondent_type = 'respondent_type' in df.columns and df['respondent_type'].notna().any()
    has_phase = PHASE_COLUMN_NAME in df.columns and df[PHASE_COLUMN_NAME].notna().any()
    
    if not has_respondent_type and not has_phase:
        st.info("資料中無身分或階段資訊，無法進行分組統計分析。")
        return
    
    # 1. 公司方 vs 投資方分析
    if has_respondent_type:
        st.markdown("#### 🏢 公司方 vs 投資方比較分析")
        
        respondent_data = df.loc[col_data.index, 'respondent_type']
        valid_types = respondent_data[respondent_data.isin(['公司方', '投資方'])]
        
        if len(valid_types.unique()) >= 2:
            if is_numeric:
                # 數值型資料：Mann-Whitney U 檢定
                company_vals = pd.to_numeric(col_data[respondent_data == '公司方'], errors='coerce').dropna()
                investor_vals = pd.to_numeric(col_data[respondent_data == '投資方'], errors='coerce').dropna()
                
                if len(company_vals) > 0 and len(investor_vals) > 0:
                    st.markdown("**描述統計：**")
                    stats_df = pd.DataFrame({
                        '群體': ['公司方', '投資方'],
                        '樣本數': [len(company_vals), len(investor_vals)],
                        '平均數': [company_vals.mean(), investor_vals.mean()],
                        '中位數': [company_vals.median(), investor_vals.median()],
                        '標準差': [company_vals.std(), investor_vals.std()],
                        '最小值': [company_vals.min(), investor_vals.min()],
                        '最大值': [company_vals.max(), investor_vals.max()]
                    })
                    st.dataframe(stats_df.style.format({
                        '平均數': '{:.2f}', '中位數': '{:.2f}', '標準差': '{:.2f}',
                        '最小值': '{:.2f}', '最大值': '{:.2f}'
                    }), use_container_width=True)
                    
                    try:
                        stat, p = mannwhitneyu(company_vals, investor_vals, alternative='two-sided')
                        st.markdown("**Mann-Whitney U 檢定結果：**")
                        st.write(f"- U 統計量 = {stat:.2f}")
                        st.write(f"- {format_p_value(p)}")
                        
                        # Cohen's d 效果量
                        pooled_std = np.sqrt(((len(company_vals)-1)*company_vals.std()**2 + (len(investor_vals)-1)*investor_vals.std()**2) / (len(company_vals)+len(investor_vals)-2))
                        cohens_d = (company_vals.mean() - investor_vals.mean()) / pooled_std if pooled_std > 0 else 0
                        st.write(f"- Cohen's d = {cohens_d:.3f} ({interpret_effect_size(cohens_d=cohens_d)})")
                        
                        st.markdown(generate_academic_conclusion(
                            test_type="mann_whitney",
                            p_value=p,
                            effect_size=cohens_d,
                            groups_info={
                                '公司方': {'n': len(company_vals), 'mean': company_vals.mean(), 'median': company_vals.median(), 'std': company_vals.std()},
                                '投資方': {'n': len(investor_vals), 'mean': investor_vals.mean(), 'median': investor_vals.median(), 'std': investor_vals.std()}
                            },
                            question_name="公司方 vs 投資方"
                        ))
                    except Exception as e:
                        st.warning(f"無法執行 Mann-Whitney U 檢定：{e}")
                else:
                    st.info("公司方或投資方的樣本數不足，無法進行統計檢定。")
            
            elif is_multiselect:
                # 複選題：對每個選項進行卡方檢定
                st.markdown("**複選題選項分析（公司方 vs 投資方）：**")
                exploded = col_data.astype(str).str.split('\n').explode().str.strip()
                exploded = exploded[(exploded != '') & (exploded != 'nan') & exploded.notna()]
                
                if not exploded.empty:
                    options = exploded.unique()
                    for opt in options[:10]:  # 限制前10個選項避免過多
                        has_opt = col_data.astype(str).apply(lambda x: opt in [s.strip() for s in str(x).split('\n') if s.strip()])
                        opt_data = pd.DataFrame({
                            'has_option': has_opt[respondent_data.isin(['公司方', '投資方'])],
                            'respondent': respondent_data[respondent_data.isin(['公司方', '投資方'])]
                        }).dropna()
                        
                        if len(opt_data) > 0:
                            table = pd.crosstab(opt_data['has_option'], opt_data['respondent'])
                            if table.shape[0] >= 2 and table.shape[1] >= 2:
                                try:
                                    chi2, p, dof, exp = chi2_contingency(table)
                                    if np.nanmin(exp) > 1:
                                        n = table.values.sum()
                                        cramers = np.sqrt(chi2 / (n * (min(table.shape) - 1)))
                                        st.write(f"**選項「{opt}」：** {format_p_value(p)}，Cramér's V = {cramers:.3f}")
                                except Exception:
                                    pass
            
            else:
                # 類別型資料：卡方檢定
                category_data = col_data[respondent_data.isin(['公司方', '投資方'])].astype(str)
                category_data = category_data[~category_data.str.lower().str.contains('nan', na=False)]
                respondent_filtered = respondent_data[category_data.index]
                
                if len(category_data) > 0:
                    table = pd.crosstab(category_data, respondent_filtered)
                    
                    st.markdown("**交叉列聯表：**")
                    st.dataframe(table, use_container_width=True)
                    
                    if table.shape[0] >= 2 and table.shape[1] >= 2:
                        try:
                            chi2, p, dof, exp = chi2_contingency(table)
                            
                            if np.nanmin(exp) > 1:
                                n = table.values.sum()
                                cramers = np.sqrt(chi2 / (n * (min(table.shape) - 1)))
                                
                                st.markdown("**卡方檢定結果：**")
                                st.write(f"- χ² = {chi2:.2f}, df = {dof}")
                                st.write(f"- {format_p_value(p)}")
                                st.write(f"- Cramér's V = {cramers:.3f} ({interpret_effect_size(cramers_v=cramers)})")
                                
                                st.markdown(generate_academic_conclusion(
                                    test_type="chi_square",
                                    p_value=p,
                                    effect_size=cramers,
                                    question_name="公司方 vs 投資方"
                                ))
                            else:
                                st.warning("期望次數過小（<1），改用 Fisher's Exact Test")
                                try:
                                    if table.shape == (2, 2):
                                        oddsratio, p = fisher_exact(table)
                                        st.write(f"- {format_p_value(p)}")
                                        st.write(f"- Odds Ratio = {oddsratio:.3f}")
                                except Exception as e:
                                    st.warning(f"無法執行 Fisher's Exact Test：{e}")
                        except Exception as e:
                            st.warning(f"無法執行卡方檢定：{e}")
        else:
            st.info("只有單一身分類型，無法進行公司方 vs 投資方比較。")
    
    # 2. 不同階段分析
    if has_phase and df[PHASE_COLUMN_NAME].nunique() > 1:
        st.markdown("#### 📊 不同階段比較分析")
        
        phase_data = df.loc[col_data.index, PHASE_COLUMN_NAME].fillna('未標註階段')
        
        if len(phase_data.unique()) >= 2:
            if is_numeric:
                # 數值型資料：Kruskal-Wallis H 檢定
                groups = []
                labels = []
                groups_info = {}
                
                for phase in sorted(phase_data.unique()):
                    phase_vals = pd.to_numeric(col_data[phase_data == phase], errors='coerce').dropna()
                    if len(phase_vals) > 0:
                        groups.append(phase_vals)
                        labels.append(phase)
                        groups_info[phase] = {
                            'n': len(phase_vals),
                            'mean': phase_vals.mean(),
                            'median': phase_vals.median(),
                            'std': phase_vals.std()
                        }
                
                if len(groups) >= 2:
                    st.markdown("**各階段描述統計：**")
                    phase_stats_df = pd.DataFrame([
                        {
                            '階段': label,
                            '樣本數': info['n'],
                            '平均數': info['mean'],
                            '中位數': info['median'],
                            '標準差': info['std']
                        }
                        for label, info in groups_info.items()
                    ])
                    st.dataframe(phase_stats_df.style.format({
                        '平均數': '{:.2f}', '中位數': '{:.2f}', '標準差': '{:.2f}'
                    }), use_container_width=True)
                    
                    try:
                        stat, p = kruskal(*groups)
                        st.markdown("**Kruskal-Wallis H 檢定結果：**")
                        st.write(f"- H 統計量 = {stat:.2f}")
                        st.write(f"- {format_p_value(p)}")
                        
                        st.markdown(generate_academic_conclusion(
                            test_type="kruskal",
                            p_value=p,
                            groups_info=groups_info,
                            question_name="階段比較"
                        ))
                    except Exception as e:
                        st.warning(f"無法執行 Kruskal-Wallis 檢定：{e}")
            
            elif is_multiselect:
                # 複選題：對每個選項進行階段間卡方檢定
                st.markdown("**複選題選項階段分析：**")
                compute_and_display_multiselect_option_tests(df, col_data, 
                    col_data.astype(str).str.split('\n').explode().str.strip().unique()[:10])
            
            else:
                # 類別型資料：卡方檢定
                category_data = col_data.astype(str)
                category_data = category_data[~category_data.str.lower().str.contains('nan', na=False)]
                phase_filtered = phase_data[category_data.index]
                
                if len(category_data) > 0:
                    table = pd.crosstab(category_data, phase_filtered)
                    
                    st.markdown("**階段交叉列聯表：**")
                    st.dataframe(table, use_container_width=True)
                    
                    if table.shape[0] >= 2 and table.shape[1] >= 2:
                        try:
                            chi2, p, dof, exp = chi2_contingency(table)
                            
                            if np.nanmin(exp) > 1:
                                n = table.values.sum()
                                cramers = np.sqrt(chi2 / (n * (min(table.shape) - 1)))
                                
                                st.markdown("**卡方檢定結果：**")
                                st.write(f"- χ² = {chi2:.2f}, df = {dof}")
                                st.write(f"- {format_p_value(p)}")
                                st.write(f"- Cramér's V = {cramers:.3f} ({interpret_effect_size(cramers_v=cramers)})")
                                
                                st.markdown(generate_academic_conclusion(
                                    test_type="chi_square",
                                    p_value=p,
                                    effect_size=cramers,
                                    question_name="階段比較"
                                ))
                            else:
                                st.warning("期望次數過小（<1），建議合併類別或增加樣本數")
                        except Exception as e:
                            st.warning(f"無法執行卡方檢定：{e}")

st.set_page_config(layout="wide", page_title="問卷互動分析報告")

@st.cache_data
def load_and_concat(file_paths):
    all_dfs = []
    for path in file_paths:
        if not isinstance(path, str) or path.strip() == "":
            continue
        if not os.path.exists(path):
            continue
        df = None
        for enc in ("utf-8", "utf-8-sig", "latin1"):
            try:
                df = pd.read_csv(path, encoding=enc)
                break
            except Exception:
                pass
        if df is None:
            continue
        try:
            df.columns = df.columns.str.replace(r'【.*?】', '', regex=True).str.strip()
            df.columns = df.columns.str.replace('\n', ' ', regex=False)
        except Exception:
            pass
        try:
            if PHASE_COLUMN_NAME in df.columns:
                extracted = df[PHASE_COLUMN_NAME].astype(str).str.extract(r'(第一階段|第二階段|第三階段)', expand=False)
                df[PHASE_COLUMN_NAME] = extracted.where(extracted.notna(), df[PHASE_COLUMN_NAME])
            else:
                m = re.search(r'(第一階段|第二階段|第三階段)', os.path.basename(path))
                if m:
                    df[PHASE_COLUMN_NAME] = m.group(1)
        except Exception:
            pass
        df['_source_file'] = os.path.basename(path)
        all_dfs.append(df)
    if not all_dfs:
        return pd.DataFrame()
    return pd.concat(all_dfs, ignore_index=True, sort=False)

st.title("📊 問卷資料互動分析報告")
st.markdown("請先選擇分析模式，然後再根據提示選擇要查看的資料範圍。")

# --- File Definitions ---
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
ALL_FILES = list(company_files.values()) + list(investor_files.values()) + [COMPANY_NEW_MULTIPHASE_FILE]

RESP_COLOR_MAP = {
    "公司方": "#1f77b4",
    "投資方": "#ff7f0e",
    "未知":   "#7f7f7f"
}

# --- UI Logic ---
analysis_mode = st.radio("**步驟一：請選擇分析模式**", ('逐題瀏覽', '合併分析'), horizontal=True, key="main_mode")

df_to_analyze = None
report_title = ""
files_to_load = []

if analysis_mode == '逐題瀏覽':
    data_source = st.radio("**步驟二：請選擇要分析的對象**", ('公司方', '投資方'), horizontal=True, key="data_source")
    files = company_files if data_source == '公司方' else investor_files
    phase_options = ["不分階段 (全部合併)"] + list(files.keys())
    selected_phase = st.radio("**步驟三：請選擇問卷階段**", phase_options, horizontal=False, key="phase_select")
    
    if selected_phase == "不分階段 (全部合併)":
        files_to_load = list(files.values())
        if data_source == '公司方':
            files_to_load.append(COMPANY_NEW_MULTIPHASE_FILE)
        report_title = f"{data_source} - 不分階段 (全部合併)"
    else:
        files_to_load = [files[selected_phase]]
        if data_source == '公司方':
            files_to_load.append(COMPANY_NEW_MULTIPHASE_FILE)
        report_title = f"{data_source} - {selected_phase}"
    
    df_to_analyze = load_and_concat(files_to_load)
    
    if selected_phase != "不分階段 (全部合併)" and data_source == '公司方':
        if PHASE_COLUMN_NAME in df_to_analyze.columns:
            df_to_analyze = df_to_analyze[
                (df_to_analyze['_source_file'].str.contains(selected_phase.replace('階段', ''), na=False)) |
                (df_to_analyze[PHASE_COLUMN_NAME].astype(str).str.contains(selected_phase, na=False))
            ]

elif analysis_mode == '合併分析':
    combine_option = st.radio("**步驟二：請選擇合併方式**", ('合併所有階段', '合併第一階段', '合併第二階段', '合併第三階段'), horizontal=False, key="combine_option")
    
    if combine_option == '合併所有階段':
        files_to_load = list(company_files.values()) + list(investor_files.values()) + [COMPANY_NEW_MULTIPHASE_FILE]
        report_title = "公司方與投資方 - 所有階段合併"
    elif combine_option == '合併第一階段':
        files_to_load = [COMPANY_P1_FILE, INVESTOR_P1_FILE, COMPANY_NEW_MULTIPHASE_FILE]
        report_title = "公司方與投資方 - 第一階段"
    elif combine_option == '合併第二階段':
        files_to_load = [COMPANY_P2_FILE, INVESTOR_P2_FILE, COMPANY_NEW_MULTIPHASE_FILE]
        report_title = "公司方與投資方 - 第二階段"
    else:
        files_to_load = [COMPANY_P3_FILE, INVESTOR_P3_FILE, COMPANY_NEW_MULTIPHASE_FILE]
        report_title = "公司方與投資方 - 第三階段"
    
    df_to_analyze = load_and_concat(files_to_load)
    
    if combine_option != '合併所有階段':
        target_phase = combine_option.replace('合併', '')
        if PHASE_COLUMN_NAME in df_to_analyze.columns:
            df_to_analyze = df_to_analyze[
                (df_to_analyze['_source_file'].str.contains(target_phase.replace('階段', ''), na=False)) |
                (df_to_analyze[PHASE_COLUMN_NAME].astype(str).str.contains(target_phase, na=False))
            ]

# 標記填答者身分
if df_to_analyze is not None and not df_to_analyze.empty:
    try:
        if '_source_file' in df_to_analyze.columns:
            def infer_role(fname):
                if not isinstance(fname, str): return '未知'
                if '投資' in fname or 'INVEST' in fname.upper():
                    return '投資方'
                return '公司方'
            df_to_analyze['respondent_type'] = df_to_analyze['_source_file'].astype(str).apply(infer_role)
        else:
            df_to_analyze['respondent_type'] = '未知'
    except Exception:
        df_to_analyze['respondent_type'] = '未知'

if df_to_analyze is None or df_to_analyze.empty:
    st.warning("在此選擇下沒有載入任何資料，請檢查您的選擇和檔案。")
    st.stop()

# --- Display Analysis ---
st.header(f"您正在查看：{report_title}的分析結果")

col_metric1, col_metric2 = st.columns(2)
with col_metric1:
    st.metric("總樣本數 (問卷份數)", len(df_to_analyze))
with col_metric2:
    if len(df_to_analyze) < 30:
        st.warning("⚠️ 樣本數 < 30，統計檢定結果可能不穩定")

cols_to_exclude = ['為了後續支付訪談費，請提供您的電子郵件地址（我們將僅用於聯繫您支付訪談費，並妥善保護您的資料）:', 'IP紀錄', '額滿結束註記', '使用者紀錄', '會員時間', 'Hash', '會員編號', '自訂ID', '備註', '填答時間', PHASE_COLUMN_NAME, '_source_file', 'respondent_type']
# 題目標準化函數
def normalize_question(q):
    """標準化題目：移除「公司」、「您投資的公司」等差異"""
    if not isinstance(q, str):
        return q
    
    # 移除「未命名題目 - 」前綴
    q = re.sub(r'^未命名題目[\s\-：:]+', '', q)
    
    # 移除常見的身分區別詞（更全面的規則）
    q = q.replace('您投資的公司有', '公司')
    q = q.replace('您投資的公司', '公司')
    q = q.replace('貴公司有', '公司')
    q = q.replace('貴公司', '公司')
    q = q.replace('公司有', '公司')
    q = q.replace('公司是否', '公司')
    q = q.replace('您認為公司', '公司')
    q = q.replace('您認為', '')
    
    # 移除題目開頭的身分前綴（包含空格、破折號、冒號等）
    q = re.sub(r'^(公司|投資方|公司方)[\s\-：:]+', '', q)
    
    # 統一「-」符號（全形、半形破折號）
    q = q.replace('－', '-').replace('—', '-').replace('–', '-')
    
    # 移除多餘空白
    q = re.sub(r'\s+', ' ', q).strip()
    
    # 移除尾部的冒號或句號
    q = q.rstrip('：:。.')
    
    return q

def normalize_question_v2(q):
    """更激進的標準化：移除所有身分標記和冗餘詞彙"""
    if not isinstance(q, str):
        return q
    
    # 0. 特殊處理：內部控制循環題目（完全統一格式）
    if '內部控制循環' in q and '建立書面控制程序與執行自評' in q:
        # 先移除項目列表 (1)(2)(3)...
        q = re.sub(r'\s*[\(（]1[\)）][^？?]*', '', q)
        # 再統一文字內容（移除標點和問號）
        q = q.replace('針對下列內部控制循環，您投資的公司在建立書面控制程序與執行自評的進度為何？', 
                     '公司針對下列內部控制循環建立書面控制程序與執行自評進度')
        q = q.replace('公司針對下列內部控制循環，建立書面控制程序與執行自評的進度為何？', 
                     '公司針對下列內部控制循環建立書面控制程序與執行自評進度')
        q = q.replace('針對下列內部控制循環，您投資的公司在建立書面控制程序與執行自評的進度為何', 
                     '公司針對下列內部控制循環建立書面控制程序與執行自評進度')
        q = q.replace('公司針對下列內部控制循環，建立書面控制程序與執行自評的進度為何', 
                     '公司針對下列內部控制循環建立書面控制程序與執行自評進度')
    
    # 1. 移除「未命名題目 - 」前綴
    q = re.sub(r'^未命名題目[\s\-－—–：:]*', '', q)
    
    # 2. 統一填空符號（先處理，避免後續被誤刪）
    q = re.sub(r'_{2,}', ' _ ', q)
    q = re.sub(r'\([\s_]*\)', ' _ ', q)
    q = re.sub(r'（[\s_]*）', ' _ ', q)
    
    # 3. 統一「董監事」相關詞彙（提前處理）
    q = q.replace('董監事席次', '董事席次')
    q = q.replace('董監事 _ 位', '董事 _ 位')
    
    # 4. 移除「在...方面」、「在...上」等介系詞片語
    q = re.sub(r'在(.{1,15}?)方面', r'\1', q)
    q = re.sub(r'在(.{1,15}?)上', r'\1', q)
    
    # 5. 統一「其」、「的」、「目前的」、「之」等語氣詞
    q = q.replace('其定期性董事會', '定期性董事會')
    q = q.replace('其董事會', '董事會')
    q = q.replace('其股東結構', '股東結構')
    q = q.replace('其董事及經理人', '董事及經理人')
    q = q.replace('其員工人數', '員工人數')
    q = q.replace('其員工分紅', '員工分紅')
    q = q.replace('的定期性董事會', '定期性董事會')
    q = q.replace('的股東結構', '股東結構')
    q = q.replace('的董事間', '董事間')
    q = q.replace('目前的董事席次', '董事席次')
    q = q.replace('目前的監察人席次', '監察人席次')
    q = q.replace('目前的', '')
    q = q.replace('之董事長', '董事長')
    q = q.replace('之董事會', '董事會')
    q = q.replace('之董事', '董事')
    q = q.replace('之監察人', '監察人')
    q = q.replace('之大股東', '大股東')
    q = q.replace('之經營團隊', '經營團隊')
    q = q.replace('之現金流量', '現金流量')
    
    # 6. 補充缺失的主題標籤
    if ' - ' not in q and '揭露董事的個別酬金' in q:
        q = '資訊透明度 - ' + q
    if ' - ' not in q and '揭露總經理及副總經理的個別酬金' in q:
        q = '資訊透明度 - ' + q
    if ' - ' not in q and '董事及經理人的酬金與公司績效連動' in q:
        q = '資訊透明度 - ' + q
    if ' - ' not in q and '諮詢顧問' in q and '頻率' in q:
        q = '董事會結構與運作 - ' + q
    if ' - ' not in q and ('董事席次' in q or '監察人席次' in q):
        q = '董事會結構與運作 - ' + q
    
    # 7. 統一身分相關詞彙（更全面的替換）
    identity_replacements = [
        # === 最高優先：精確完整匹配（包含所有可能的變體）===
        # 內部控制循環題目（特殊處理：投資方版本缺少項目列表）
        ('針對下列內部控制循環，您投資的公司在建立書面控制程序與執行自評的進度為何？', '公司針對下列內部控制循環建立書面控制程序與執行自評進度'),
        # === 以上為新增 ===
        # 特定句型優先處理（更詳細的對應）
        ('您主要投資的未上市（櫃）公司所屬產業類別', '主要產業類別'),
        ('您投資的公司其員工人數', '員工人數'),
        ('您投資的公司其員工分紅', '公司員工分紅'),
        ('您投資的公司其股東結構中包含法人股東（如創投）', '公司股東結構中包含法人股東'),
        ('公司的股東結構中包含法人股東或創投', '公司股東結構中包含法人股東'),
        ('您投資的公司在現金流量規劃與監控制度的建立程度如何', '公司現金流量規劃與監控制度建立程度'),
        ('您認為公司現金流量規劃與監控制度的建立程度如何', '公司現金流量規劃與監控制度建立程度'),
        ('您投資的公司在建立書面核准流程有困難', '公司建立書面核准流程是挑戰'),
        ('建立書面核准流程對公司來說是一項挑戰', '公司建立書面核准流程是挑戰'),
        ('承上題，您投資的公司之現金流量足以支撐公司營運幾個月', '承上題公司現金流量足以支撐公司營運幾個月'),
        ('承上題您認為公司現金流量足以支撐公司營運幾個月', '承上題公司現金流量足以支撐公司營運幾個月'),
        ('您投資的公司有清楚的向股東揭露董事的個別酬金', '公司清楚的向股東揭露董事的個別酬金'),
        ('您投資的公司有清楚的向股東揭露總經理及副總經理的個別酬金', '公司清楚的向股東揭露總經理及副總經理的個別酬金'),
        ('請問您投資的公司之大股東（持股5%以上）人數有多少人', '公司大股東（持股5%以上）人數'),
        ('請問公司的大股東（持股5%以上）人數多少人', '公司大股東（持股5%以上）人數'),
        ('請問您投資的公司之大股東', '公司大股東'),
        ('請問您投資的公司', '公司'),
        ('您投資的公司在過去12個月內，董事會的召開頻率為何', '公司過去12個月內，董事會召開頻率'),
        ('在過去12個月內，貴公司董事會的召開頻率為何', '公司過去12個月內，董事會召開頻率'),
        ('您投資的公司', '公司'),
        ('請填寫公司董事席次', '公司董事席次'),
        ('請填寫公司監察人席次', '公司監察人席次'),
        ('請填寫公司董監事席次', '公司董事席次'),
        ('您投資的公司其定期性董事會的議事內容', '公司定期性董事會的議事內容'),
        ('您投資的公司定期性董事會的議事內容', '公司定期性董事會的議事內容'),
        ('公司定期性董事會的議事內容', '公司定期性董事會的議事內容'),
        ('您投資的公司有清楚的向股東揭露', '公司清楚的向股東揭露'),
        ('您投資的公司清楚的向股東揭露', '公司清楚的向股東揭露'),
        ('貴公司有清楚的向股東揭露', '公司清楚的向股東揭露'),
        ('貴公司清楚的向股東揭露', '公司清楚的向股東揭露'),
        ('您投資的公司在諮詢顧問', '公司諮詢顧問'),
        ('您投資的公司諮詢顧問', '公司諮詢顧問'),
        ('您投資的公司在訂定財會作業程序上會', '公司訂定財會作業程序'),
        ('訂定財會作業程序對公司來說', '公司訂定財會作業程序'),
        ('請填寫公司', '公司'),
        ('貴公司董事會', '公司董事會'),
        ('貴公司', '公司'),
        # 通用替換
        ('您投資的公司有', '公司'),
        ('您投資的公司其', '公司'),
        ('您投資的公司在', '公司'),
        ('您投資的公司會', '公司'),
        ('貴公司有', '公司'),
        ('貴公司在', '公司'),
        ('您認為公司', '公司'),
        ('您認為', ''),
        ('請問公司', '公司'),
        ('請填寫', ''),
    ]
    
    for old, new in identity_replacements:
        q = q.replace(old, new)
    
    # 8. 移除題目開頭的冗餘前綴（修正正則表達式）
    q = re.sub(r'^(公司方[\s\-－—–：:]+|投資方[\s\-－—–：:]+|請問[\s\-－—–：:]*|請填寫[\s\-－—–：:]*)', '', q)
    
    # 9. 統一冒號和「位」的格式
    q = q.replace('： 董事', '：董事')
    q = q.replace(': 董事', '：董事')
    q = q.replace('： 監察人', '：監察人')
    q = q.replace(': 監察人', '：監察人')
    q = re.sub(r'：[\s]+董事', '：董事', q)
    q = re.sub(r'：[\s]+監察人', '：監察人', q)
    
    # 10. 統一「頻率為何」、「為何」、「如何」、「多少人」等問句
    q = q.replace('的召開頻率為何', '召開頻率')
    q = q.replace('召開頻率為何', '召開頻率')
    q = q.replace('的頻率為何？', '頻率')
    q = q.replace('頻率為何？', '頻率')
    q = q.replace('為何？', '')
    q = q.replace('如何？', '')
    q = q.replace('的頻率', '頻率')
    q = q.replace('的建立程度如何', '建立程度')
    q = q.replace('建立程度如何', '建立程度')
    q = q.replace('的進度為何', '進度')
    q = q.replace('進度為何', '進度')
    q = q.replace('人數有多少人', '人數')
    q = q.replace('人數多少人', '人數')
    q = q.replace('有多少人', '')
    q = q.replace('多少人', '')
    
    # 11. 統一標點符號
    q = q.replace('－', ' - ').replace('—', ' - ').replace('–', ' - ')
    q = q.replace('：', ':').replace('。', '.')
    q = q.replace('？', '').replace('?', '')
    
    # 12. 統一括號與複選標記
    q = q.replace('(可複選)', '').replace('（可複選）', '')
    q = q.replace('(複選)', '').replace('（複選）', '')
    q = q.replace('（如創投）', '')
    q = q.replace('或創投', '')
    
    # 移除括號內的詳細說明（包含多個空格的情況）
    q = re.sub(r'\s{2,}\([^\)]+\)', '', q)
    q = re.sub(r'\s*\([^\)]{10,}\)', '', q)
    q = re.sub(r'\s*（[^）]{10,}）', '', q)
    
    # 13. 移除「位」前的多餘空格和符號
    q = re.sub(r'[\s_]+位', '位', q)
    
    # 14. 統一「是/會/有」等助動詞和語氣詞
    q = q.replace('來說是', '')
    q = q.replace('對公司來說是一項挑戰', '是挑戰')
    q = q.replace('對公司來說是不小的負擔', '是負擔')
    q = q.replace('上會是', '')
    q = q.replace('會是', '')
    q = q.replace('有困難', '是挑戰')
    q = q.replace('是不小的負擔', '是負擔')
    
    # 15. 移除多餘空白
    q = re.sub(r'\s+', ' ', q).strip()
    
    # 16. 移除尾部標點
    q = q.rstrip('：:。.,;；？?')
    
    return q

def calculate_similarity(s1, s2):
    """計算兩個字串的相似度 (0-1)，考慮核心內容差異"""
    # 使用 SequenceMatcher 計算基礎相似度
    base_similarity = SequenceMatcher(None, s1, s2).ratio()
    
    # 如果相似度很高，進一步檢查關鍵詞差異
    if base_similarity > 0.8:
        # 提取關鍵名詞（避免誤合併不同主題的題目）
        keywords_s1 = set(re.findall(r'[\u4e00-\u9fff]{2,}', s1))
        keywords_s2 = set(re.findall(r'[\u4e00-\u9fff]{2,}', s2))
        
        # 計算關鍵詞交集比例
        if keywords_s1 and keywords_s2:
            keyword_overlap = len(keywords_s1 & keywords_s2) / max(len(keywords_s1), len(keywords_s2))
            # 調整相似度：如果關鍵詞差異大，降低相似度
            return base_similarity * (0.5 + 0.5 * keyword_overlap)
    
    return base_similarity

def merge_similar_questions(df, cols_to_exclude, similarity_threshold=0.75):  # 降低到 0.75
    """
    基於相似度合併題目（更積極處理「未命名題目」與單方題目）
    
    Returns:
        - merged_mapping: {代表題目: [所有原始題目]}
        - cols_to_analyze: 去重後的題目列表
    """
    all_cols = [c for c in df.columns if c not in cols_to_exclude]
    
    # 第一步：標準化並分組（標準化後相同的題目會自動合併）
    normalized_groups = {}
    for col in all_cols:
        norm = normalize_question_v2(col)
        if norm not in normalized_groups:
            normalized_groups[norm] = []
        normalized_groups[norm].append(col)
    
    # 第二步：相似度匹配（處理標準化後仍有細微差異的情況）
    merged_mapping = {}
    processed = set()
    
    norm_keys = list(normalized_groups.keys())
    for i, norm1 in enumerate(norm_keys):
        if norm1 in processed:
            continue
        
        # 找出所有相似的標準化題目（包括 norm1 本身）
        similar_group = [norm1]
        for norm2 in norm_keys[i+1:]:
            if norm2 in processed:
                continue
            similarity = calculate_similarity(norm1, norm2)
            if similarity >= similarity_threshold:
                similar_group.append(norm2)
                processed.add(norm2)
        
        # 合併所有相似題目的原始欄位
        all_originals = []
        for norm in similar_group:
            all_originals.extend(normalized_groups[norm])
        
        # 優先選擇沒有「未命名題目」且較短的作為代表（公司方優先）
        representative = None
        for orig in sorted(all_originals, key=lambda x: (len(x), '投資' in x)):
            if '未命名題目' not in orig:
                representative = orig
                break
        if representative is None:  # 如果全部都是未命名題目
            representative = all_originals[0]
        
        merged_mapping[representative] = all_originals
        processed.add(norm1)
    
    # 第三步：資料合併
    for representative, originals in merged_mapping.items():
        if len(originals) > 1:
            for other_col in originals[1:]:
                # 優先保留代表題目的資料，用其他題目填補缺失
                mask = df[representative].isna() & df[other_col].notna()
                df.loc[mask, representative] = df.loc[mask, other_col]

    cols_to_analyze = list(merged_mapping.keys())
    return merged_mapping, cols_to_analyze

def generate_report_recommendations(df, cols_to_analyze, analysis_mode):
    """分析並推薦值得納入報告的題目"""
    recommendations = []
    processed_cols = set()
    
    for col_name in cols_to_analyze:
        if col_name not in df.columns or col_name in processed_cols:
            continue
        
        processed_cols.add(col_name)
        col_series = df[col_name].dropna()
        if col_series.empty or len(col_series) < 5:
            continue
        
        recommendation = {
            '題目': col_name[:80] + '...' if len(col_name) > 80 else col_name,
            '完整題目': col_name,
            '樣本數': int(df[col_name].notna().sum()),
            '缺失率': f"{(df[col_name].isna().sum() / len(df) * 100):.1f}%",
            '推薦理由': [],
            '優先順序': 0.0,
            '統計結果': {}
        }
        
        is_multiselect = col_series.dtype == 'object' and col_series.astype(str).str.contains('\n', na=False).any()
        
        # 只在合併分析且有 respondent_type 時進行比較檢定
        if analysis_mode == '合併分析' and 'respondent_type' in df.columns:
            try:
                if is_multiselect:
                    exploded = col_series.astype(str).str.split('\n').explode().str.strip()
                    exploded = exploded[(exploded != '') & (exploded != 'nan') & exploded.notna()]
                    if not exploded.empty:
                        total_counts = exploded.value_counts()
                        significant_count = 0
                        for opt in total_counts.index[:10]:
                            if pd.isna(opt) or str(opt).lower() == 'nan':
                                continue
                            pres = df[col_name].astype(str).fillna('').apply(
                                lambda s: opt in [x.strip() for x in s.split('\n') if x.strip()!='' and x.strip().lower()!='nan']
                            )
                            table = pd.crosstab(pres, df['respondent_type'])
                            if table.size > 0 and table.values.sum() > 0 and table.shape[0] >= 2:
                                try:
                                    chi2, p, dof, exp = chi2_contingency(table)
                                    if np.nanmin(exp) > 1 and p < 0.05:
                                        significant_count += 1
                                        recommendation['統計結果'].setdefault('顯著選項', []).append({'選項': opt, 'p': p})
                                        if p < 0.001:
                                            recommendation['優先順序'] += 3
                                        elif p < 0.01:
                                            recommendation['優先順序'] += 2
                                        else:
                                            recommendation['優先順序'] += 1
                                except Exception:
                                    pass
                        if significant_count > 0:
                            recommendation['推薦理由'].append(f"有 {significant_count} 個選項在公司方/投資方間呈現統計顯著差異")
                            recommendation['統計結果']['顯著選項數'] = significant_count
                else:
                    is_numeric = pd.api.types.is_numeric_dtype(col_series)
                    if not is_numeric:
                        numeric_version = pd.to_numeric(col_series, errors='coerce').dropna()
                        if len(numeric_version) > 0 and (len(numeric_version) / len(col_series) > 0.7):
                            is_numeric = True
                            col_num = numeric_version
                        else:
                            is_numeric = False
                    else:
                        col_num = pd.to_numeric(col_series, errors='coerce').dropna()
                    
                    if is_numeric:
                        groups = []
                        for rt in df['respondent_type'].unique():
                            grp = col_num[df.loc[col_num.index, 'respondent_type'] == rt]
                            if len(grp) > 0:
                                groups.append(grp.astype(float))
                        if len(groups) == 2:
                            try:
                                stat, p = mannwhitneyu(groups[0], groups[1], alternative='two-sided')
                                median_diff = abs(np.median(groups[0]) - np.median(groups[1]))
                                recommendation['統計結果']['p'] = float(p)
                                recommendation['統計結果']['median_diff'] = float(median_diff)
                                if p < 0.05:
                                    recommendation['推薦理由'].append(f"公司方/投資方中位數差異顯著 (p={p:.3f})")
                                    recommendation['優先順序'] += 2
                            except Exception:
                                pass
                        elif len(groups) > 2:
                            try:
                                stat, p = kruskal(*groups)
                                if p < 0.05:
                                    recommendation['推薦理由'].append("跨組差異顯著 (Kruskal-Wallis)")
                                    recommendation['優先順序'] += 2
                            except Exception:
                                pass
                    else:
                        s = col_series.astype(str)
                        s = s[~s.str.lower().str.contains('nan', na=False)]
                        if not s.empty:
                            table = pd.crosstab(s, df.loc[s.index, 'respondent_type'])
                            if table.size > 0 and table.values.sum() > 0:
                                try:
                                    if table.shape == (2, 2) and table.values.sum() < 20:
                                        oddsratio, p = fisher_exact(table)
                                    else:
                                        chi2, p, dof, exp = chi2_contingency(table)
                                    
                                    if p < 0.05:
                                        recommendation['推薦理由'].append(f"公司方/投資方分佈顯著差異 (p={p:.3f})")
                                        recommendation['統計結果']['p'] = float(p)
                                        if p < 0.001:
                                            recommendation['優先順序'] += 3
                                        elif p < 0.01:
                                            recommendation['優先順序'] += 2
                                        else:
                                            recommendation['優先順序'] += 1
                                except Exception:
                                    pass
            except Exception:
                pass
        
        # 額外評分標準
        missing_rate = df[col_name].isna().sum() / len(df)
        if missing_rate < 0.05:
            recommendation['推薦理由'].append("資料完整度高 (缺失 < 5%)")
            recommendation['優先順序'] += 1
        
        if not is_multiselect and not col_series.empty:
            unique_ratio = len(col_series.unique()) / len(col_series)
            if unique_ratio > 0.3:
                recommendation['推薦理由'].append("答案具多樣性")
                recommendation['優先順序'] += 0.5
        
        if recommendation['推薦理由']:
            recommendations.append(recommendation)
    
    recommendations.sort(key=lambda x: x['優先順序'], reverse=True)
    return recommendations

def generate_professional_report(df, recommendations, cols_to_analyze, analysis_mode):
    """
    生成符合國發基金需求的專業分析報告
    結構：執行摘要 → 方法論 → 主要發現 → 結論與建議
    """
    report = []
    
    # === 1. 標題與基本資訊 ===
    report.append("# 未上市櫃公司治理問卷分析報告")
    report.append(f"\n**報告產生時間：** {datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
    report.append(f"\n**分析模式：** {analysis_mode}")
    report.append(f"\n**總樣本數：** {len(df)} 筆")
    
    if 'respondent_type' in df.columns:
        respondent_counts = df['respondent_type'].value_counts()
        report.append(f"\n**填答者分佈：**")
        for resp_type, count in respondent_counts.items():
            report.append(f"- {resp_type}：{count} 筆 ({count/len(df)*100:.1f}%)")
    
    if 'phase' in df.columns and df['phase'].notna().any():
        phase_counts = df['phase'].value_counts()
        report.append(f"\n**階段分佈：**")
        for phase, count in phase_counts.items():
            report.append(f"- {phase}：{count} 筆 ({count/len(df)*100:.1f}%)")
    
    report.append("\n---\n")
    
    # === 2. 執行摘要 ===
    report.append("## 📋 執行摘要\n")
    report.append("本報告針對未上市櫃公司治理問卷進行全面性統計分析，主要目的在於瞭解公司方與投資方對公司治理實務的認知差異，以及不同階段公司在治理面向的發展狀況。\n")
    
    # 找出最重要的3-5個發現
    top_findings = recommendations[:min(5, len(recommendations))]
    report.append("### 關鍵發現：\n")
    for idx, rec in enumerate(top_findings, 1):
        topic = rec['完整題目']
        priority = rec['優先順序']
        reasons = rec['推薦理由']
        
        # 將統計術語轉為業務語言
        business_insight = []
        for reason in reasons:
            if "公司方/投資方" in reason and "顯著差異" in reason:
                business_insight.append("**公司方與投資方對此議題的看法存在顯著落差**，建議關注雙方認知差異的根源")
            elif "分佈顯著差異" in reason:
                business_insight.append("**不同群體在此議題上呈現明顯差異**，值得進一步探討造成差異的因素")
            elif "資料完整度高" in reason:
                business_insight.append("此議題獲得高度關注，資料品質優良")
            elif "答案具多樣性" in reason:
                business_insight.append("受訪者回應具多樣性，反映實務做法的多元性")
        
        report.append(f"{idx}. **{topic[:60]}{'...' if len(topic) > 60 else ''}**")
        report.append(f"   - 重要性評分：{priority:.1f} 分")
        if business_insight:
            report.append(f"   - 業務意涵：{business_insight[0]}")
        report.append("")
    
    report.append("\n---\n")
    
    # === 3. 方法論 ===
    report.append("## 🔬 研究方法論\n")
    report.append("### 3.1 資料來源與樣本\n")
    report.append(f"本研究分析 {len(df)} 筆問卷資料，涵蓋 {len(cols_to_analyze)} 個分析面向。")
    
    if 'respondent_type' in df.columns:
        report.append("資料來源包含公司方填答與投資方填答，可進行雙向比對分析。\n")
    
    report.append("### 3.2 統計分析方法\n")
    report.append("本研究採用以下統計方法：\n")
    report.append("1. **描述性統計**：計算次數分佈、百分比、平均數、中位數等基本統計量")
    report.append("2. **卡方檢定（Chi-square test）**：檢驗類別變項在不同群體間的分佈差異")
    report.append("3. **Mann-Whitney U 檢定**：檢驗數值變項在兩組間的分佈差異（非參數檢定）")
    report.append("4. **Kruskal-Wallis 檢定**：檢驗數值變項在多組間的分佈差異（非參數檢定）")
    report.append("5. **Fisher 精確檢定**：針對小樣本的類別變項進行精確機率檢定\n")
    
    report.append("### 3.3 顯著性水準\n")
    report.append("本研究採用以下顯著性標準：")
    report.append("- p < 0.001：極顯著差異 (⭐⭐⭐)")
    report.append("- p < 0.01：非常顯著差異 (⭐⭐)")
    report.append("- p < 0.05：顯著差異 (⭐)")
    report.append("- p ≥ 0.05：無顯著差異\n")
    
    report.append("\n---\n")
    
    # === 4. 主要發現 ===
    report.append("## 📊 主要發現\n")
    
    # 按優先順序分組
    high_priority = [r for r in recommendations if r['優先順序'] >= 3]
    medium_priority = [r for r in recommendations if 2 <= r['優先順序'] < 3]
    
    if high_priority:
        report.append("### 4.1 高度關注議題（優先順序 ≥ 3）\n")
        report.append("以下議題在統計分析中呈現極顯著或多重顯著差異，建議優先關注：\n")
        
        for idx, rec in enumerate(high_priority, 1):
            report.append(f"#### 議題 {idx}：{rec['完整題目']}\n")
            report.append(f"**樣本數：** {rec['樣本數']} | **缺失率：** {rec['缺失率']} | **優先順序：** {rec['優先順序']:.1f}\n")
            
            # 統計結果解讀
            if '統計結果' in rec and rec['統計結果']:
                stats = rec['統計結果']
                
                if 'p' in stats:
                    p_val = stats['p']
                    sig_level = "極顯著" if p_val < 0.001 else "非常顯著" if p_val < 0.01 else "顯著"
                    report.append(f"**統計檢定結果：**")
                    report.append(f"- p-value = {p_val:.4f} ({sig_level})")
                    
                    if 'median_diff' in stats:
                        report.append(f"- 中位數差異：{stats['median_diff']:.2f}")
                    
                    # 業務解讀
                    report.append(f"\n**業務解讀：**")
                    if p_val < 0.001:
                        report.append("此議題在不同群體間存在極顯著差異（p < 0.001），顯示雙方在認知或實務上有本質性的差距。建議深入探討造成差異的結構性因素，並評估是否需要政策介入或輔導機制。")
                    elif p_val < 0.01:
                        report.append("此議題呈現高度顯著差異（p < 0.01），反映不同群體在此面向的經驗或期待有明顯落差。建議納入後續輔導計畫的重點項目。")
                    else:
                        report.append("此議題存在顯著差異（p < 0.05），值得關注並進一步分析差異成因。")
                
                if '顯著選項數' in stats:
                    sig_count = stats['顯著選項數']
                    report.append(f"\n- 有 {sig_count} 個選項呈現顯著差異")
                    report.append(f"- **解讀：** 此複選題中有多個選項在不同群體間分佈不均，顯示在具體實務做法上存在系統性差異。")
            
            report.append("\n" + "- " * 30 + "\n")
    
    if medium_priority:
        report.append("\n### 4.2 重要議題（優先順序 2-3）\n")
        report.append("以下議題具有統計顯著性或高資料完整度，值得納入報告：\n")
        
        for idx, rec in enumerate(medium_priority, 1):
            report.append(f"**{idx}. {rec['完整題目'][:80]}{'...' if len(rec['完整題目']) > 80 else ''}**")
            report.append(f"- 樣本數：{rec['樣本數']} | 缺失率：{rec['缺失率']}")
            report.append(f"- 重點：{'; '.join(rec['推薦理由'][:2])}")
            report.append("")
    
    report.append("\n---\n")
    
    # === 5. 結論與建議 ===
    report.append("## 💡 結論與政策建議\n")
    
    report.append("### 5.1 總體觀察\n")
    report.append(f"本次問卷分析涵蓋 {len(recommendations)} 個具有分析價值的議題，")
    report.append(f"其中 {len(high_priority)} 個議題呈現高度顯著差異，{len(medium_priority)} 個議題具有重要參考價值。\n")
    
    if 'respondent_type' in df.columns:
        report.append("### 5.2 公司方與投資方的認知落差\n")
        report.append("分析顯示公司方與投資方在多項公司治理議題上存在認知或實務差異。")
        report.append("此落差可能來自於：")
        report.append("- **資訊不對稱**：投資方對公司實務的了解程度有限")
        report.append("- **期待差異**：雙方對治理標準的認知不一致")
        report.append("- **實務落差**：公司自評與外部評估的客觀性差異\n")
    
    report.append("### 5.3 政策建議\n")
    report.append("基於上述分析結果，本研究提出以下政策建議供國發基金參考：\n")
    
    # 根據高優先順序議題生成具體建議
    if high_priority:
        report.append("**針對高度關注議題：**\n")
        
        # 分析是否有特定領域的問題
        governance_issues = [r for r in high_priority if any(kw in r['完整題目'] for kw in ['董事會', '董事', '監察人'])]
        transparency_issues = [r for r in high_priority if any(kw in r['完整題目'] for kw in ['揭露', '透明', '資訊'])]
        internal_control_issues = [r for r in high_priority if any(kw in r['完整題目'] for kw in ['內部控制', '流程', '制度'])]
        
        if governance_issues:
            report.append("1. **強化董事會運作機制**")
            report.append("   - 建議提供未上市櫃公司治理訓練課程")
            report.append("   - 推動獨立董事或外部董事制度")
            report.append("   - 建立董事會運作評估機制\n")
        
        if transparency_issues:
            report.append("2. **提升資訊透明度**")
            report.append("   - 建立資訊揭露標準範本")
            report.append("   - 鼓勵定期向股東報告")
            report.append("   - 推動數位化資訊平台\n")
        
        if internal_control_issues:
            report.append("3. **建立內部控制制度**")
            report.append("   - 提供內控建置輔導服務")
            report.append("   - 分享最佳實務案例")
            report.append("   - 建立分階段導入機制\n")
    
    report.append("4. **縮小公司方與投資方認知落差**")
    report.append("   - 定期舉辦溝通座談會")
    report.append("   - 建立雙向回饋機制")
    report.append("   - 提供第三方治理評估服務\n")
    
    report.append("5. **階段性輔導機制**")
    report.append("   - 針對不同發展階段提供客製化輔導")
    report.append("   - 建立標竿企業示範案例")
    report.append("   - 提供持續追蹤與評估\n")
    
    report.append("\n---\n")
    
    # === 6. 附錄 ===
    report.append("## 📎 附錄\n")
    report.append("### 附錄 A：完整分析議題清單\n")
    report.append(f"本次分析共涵蓋 {len(recommendations)} 個議題，完整清單如下：\n")
    
    report.append("| 排名 | 題目 | 樣本數 | 缺失率 | 優先順序 |")
    report.append("|------|------|--------|--------|----------|")
    
    for idx, rec in enumerate(recommendations[:20], 1):  # 只顯示前20題
        topic_short = rec['題目'][:40] + '...' if len(rec['題目']) > 40 else rec['題目']
        report.append(f"| {idx} | {topic_short} | {rec['樣本數']} | {rec['缺失率']} | {rec['優先順序']:.1f} |")
    
    if len(recommendations) > 20:
        report.append(f"\n*註：完整清單包含 {len(recommendations)} 個議題，此處僅顯示前 20 題*\n")
    
    report.append("\n### 附錄 B：統計方法說明\n")
    report.append("**卡方檢定（Chi-square test）**")
    report.append("- 適用於類別變項的獨立性檢定")
    report.append("- 零假設：兩個類別變項之間獨立（無關聯）")
    report.append("- 當 p < 0.05 時拒絕零假設，認為變項間存在關聯\n")
    
    report.append("**Mann-Whitney U 檢定**")
    report.append("- 非參數檢定方法，不假設資料符合常態分佈")
    report.append("- 適用於比較兩組獨立樣本的分佈")
    report.append("- 檢驗兩組的中位數是否有顯著差異\n")
    
    report.append("**Kruskal-Wallis 檢定**")
    report.append("- Mann-Whitney U 檢定的擴展版本")
    report.append("- 適用於比較三組或以上獨立樣本")
    report.append("- 檢驗多組間是否存在顯著差異\n")
    
    report.append("\n---\n")
    report.append(f"\n**報告結束** | 產生時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    return "\n".join(report)

# 執行題目合併
st.markdown("### 🔄 正在進行題目去重與合併...")
with st.spinner("分析題目相似度中..."):
    if analysis_mode == '合併分析':
        merged_mapping, cols_to_analyze = merge_similar_questions(
            df_to_analyze, 
            cols_to_exclude, 
            similarity_threshold=0.70  # 降低閾值，更積極合併
        )
    else:
        # 逐題瀏覽模式：不合併，直接使用所有欄位
        cols_to_analyze = [c for c in df_to_analyze.columns if c not in cols_to_exclude]
        merged_mapping = {c: [c] for c in cols_to_analyze}  # 建立一對一映射

# 顯示合併結果（只在合併分析模式下顯示）
if analysis_mode == '合併分析':
    with st.expander("🔍 題目合併詳細資訊（除錯用）", expanded=False):
        duplicate_groups = {k: v for k, v in merged_mapping.items() if len(v) > 1}
        
        if duplicate_groups:
            st.success(f"✅ 成功合併 {len(duplicate_groups)} 組重複題目，共減少 {sum(len(v)-1 for v in duplicate_groups.values())} 個重複項")
            
            # 統計合併效果
            if 'respondent_type' in df_to_analyze.columns:
                company_only = 0
                investor_only = 0
                mixed = 0
                
                for representative, originals in duplicate_groups.items():
                    respondent_types = set()
                    for orig in originals:
                        data = df_to_analyze[orig].dropna()
                        if not data.empty:
                            types = df_to_analyze.loc[data.index, 'respondent_type'].unique()
                            respondent_types.update(types)
                    
                    if '公司方' in respondent_types and '投資方' in respondent_types:
                        mixed += 1
                    elif '公司方' in respondent_types:
                        company_only += 1
                    elif '投資方' in respondent_types:
                        investor_only += 1
                
                st.write(f"- 🔵 公司方專用題目合併：{company_only} 組")
                st.write(f"- 🟠 投資方專用題目合併：{investor_only} 組")
                st.write(f"- 🟢 跨身分題目合併：{mixed} 組")
            
            # 顯示範例（前 10 組）
            st.markdown("**合併範例（前 10 組）：**")
            for i, (representative, originals) in enumerate(list(duplicate_groups.items())[:10], 1):
                st.markdown(f"**{i}. 代表題目：** {representative}")
                normalized_rep = normalize_question_v2(representative)
                st.caption(f"標準化為：{normalized_rep}")
                
                for orig in originals:
                    if orig == representative:
                        continue
                    similarity = calculate_similarity(
                        normalize_question_v2(representative), 
                        normalize_question_v2(orig)
                    )
                    orig_data = df_to_analyze[orig].dropna()
                    if not orig_data.empty and 'respondent_type' in df_to_analyze.columns:
                        respondents = df_to_analyze.loc[orig_data.index, 'respondent_type'].value_counts().to_dict()
                        resp_str = ", ".join([f"{k}:{v}筆" for k, v in respondents.items()])
                        st.write(f"  ↳ {orig}")
                        st.caption(f"    相似度: {similarity:.2%} | 資料: {resp_str}")
                    else:
                        st.write(f"  ↳ {orig} (無資料)")
                
                st.markdown("---")
            
            if len(duplicate_groups) > 10:
                st.info(f"還有 {len(duplicate_groups)-10} 組合併題目未顯示...")
        else:
            st.success("✅ 沒有發現需要合併的重複題目")
        
        st.metric("最終分析題目數", len(cols_to_analyze), 
                  delta=f"-{len(df_to_analyze.columns) - len(cols_to_exclude) - len(cols_to_analyze)}" if len(df_to_analyze.columns) - len(cols_to_exclude) > len(cols_to_analyze) else "0")

# --- 功能區（保留原有功能）---
st.markdown("---")

# 生成報告推薦
if analysis_mode == '合併分析':
    st.markdown("---")
    st.subheader("📋 適合寫入報告的題目推薦")
    
    with st.spinner("正在分析並推薦重要題目..."):
        recommendations = generate_report_recommendations(df_to_analyze, cols_to_analyze, analysis_mode)
    
    if recommendations:
        st.success(f"✅ 找到 {len(recommendations)} 題具有分析價值的題目")
        
        # 顯示前 20 題推薦
        rec_df = pd.DataFrame([{
            '排名': i+1,
            '題目': rec['題目'],
            '樣本數': rec['樣本數'],
            '缺失率': rec['缺失率'],
            '推薦理由': '；'.join(rec['推薦理由']),
            '優先順序分數': f"{rec['優先順序']:.1f}"
        } for i, rec in enumerate(recommendations[:20])])
        
        st.info("💡 **使用建議**：優先順序分數 ≥ 2 的題目通常具有較高的報告價值")
        st.dataframe(rec_df, use_container_width=True)
        
        # === 新增：深度分析報告 ===
        st.markdown("---")
        st.markdown("### 📊 深度分析報告")
        
        # 讓使用者選擇要深入分析的題目
        high_priority_recs = [rec for rec in recommendations if rec['優先順序'] >= 2]
        if high_priority_recs:
            selected_topics = st.multiselect(
                "選擇要深入分析的題目（預設為優先順序 ≥ 2 的題目）:",
                options=[rec['完整題目'] for rec in high_priority_recs],
                default=[rec['完整題目'] for rec in high_priority_recs[:5]]  # 預設前5題
            )
            
            if selected_topics:
                for topic in selected_topics:
                    # 找到對應的推薦資訊
                    rec_info = next((r for r in recommendations if r['完整題目'] == topic), None)
                    if not rec_info:
                        continue
                    
                    with st.expander(f"📈 {rec_info['題目']}", expanded=False):
                        col_data = df_to_analyze[topic].dropna()
                        if col_data.empty:
                            st.warning("無有效資料")
                            continue
                        
                        # 顯示統計摘要
                        st.markdown("#### 📋 基本資訊")
                        info_cols = st.columns(3)
                        info_cols[0].metric("樣本數", rec_info['樣本數'])
                        info_cols[1].metric("缺失率", rec_info['缺失率'])
                        info_cols[2].metric("優先順序", f"{rec_info['優先順序']:.1f}")
                        
                        st.markdown("**推薦理由：**")
                        for reason in rec_info['推薦理由']:
                            st.write(f"- {reason}")
                        
                        # 判斷題型
                        is_multiselect = col_data.dtype == 'object' and col_data.astype(str).str.contains('\n', na=False).any()
                        is_numeric = pd.api.types.is_numeric_dtype(col_data)
                        
                        # 統一處理數值資料
                        col_data_numeric = None
                        if is_numeric:
                            col_data_numeric = pd.to_numeric(col_data, errors='coerce').dropna()
                        else:
                            numeric_version = pd.to_numeric(col_data, errors='coerce').dropna()
                            if len(numeric_version) > 0 and (len(numeric_version) / len(col_data) > 0.7):
                                is_numeric = True
                                col_data_numeric = numeric_version
                        
                        # === 分析1: 公司方 vs 投資方 ===
                        if 'respondent_type' in df_to_analyze.columns:
                            st.markdown("---")
                            st.markdown("#### 🔵🟠 公司方 vs 投資方比較")
                            
                            if is_multiselect:
                                # 複選題分析
                                exploded = col_data.astype(str).str.split('\n').explode().str.strip()
                                exploded = exploded[(exploded != '') & (exploded != 'nan') & exploded.notna()]
                                
                                if not exploded.empty:
                                    df_exp = exploded.to_frame(name='option')
                                    df_exp['respondent_type'] = df_to_analyze.loc[df_exp.index, 'respondent_type'].fillna('未知')
                                    
                                    # 計算各選項在不同身分的比例
                                    crosstab = pd.crosstab(df_exp['option'], df_exp['respondent_type'], normalize='columns') * 100
                                    
                                    # 智慧排序 x 軸
                                    sorted_index = smart_sort_categories(crosstab.index)
                                    crosstab = crosstab.reindex(sorted_index)
                                    
                                    if crosstab.shape[1] >= 2:
                                        # 繪製堆疊長條圖
                                        fig = go.Figure()
                                        colors = {'公司方': '#1f77b4', '投資方': '#ff7f0e', '未知': '#999999'}
                                        
                                        for resp_type in crosstab.columns:
                                            fig.add_trace(go.Bar(
                                                name=resp_type,
                                                x=crosstab.index,
                                                y=crosstab[resp_type],
                                                marker_color=colors.get(resp_type, '#cccccc'),
                                                text=[f"{v:.1f}%" for v in crosstab[resp_type]],
                                                textposition='auto'
                                            ))
                                        
                                        fig.update_layout(
                                            barmode='group',
                                            title='各選項在不同身分的選擇比例',
                                            xaxis_title='選項',
                                            yaxis_title='比例 (%)',
                                            template='plotly_white',
                                            height=500,
                                            xaxis_tickangle=-45,
                                            xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                                        )
                                        st.plotly_chart(fig, use_container_width=True)
                                        
                                        # 顯著差異的選項
                                        if '顯著選項' in rec_info['統計結果']:
                                            st.markdown("**統計檢定結果（卡方檢定）：**")
                                            for sig_opt in rec_info['統計結果']['顯著選項'][:5]:
                                                p_val = sig_opt['p']
                                                significance = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*"
                                                st.write(f"- 選項「{sig_opt['選項']}」：公司方與投資方選擇比例有顯著差異 (p = {p_val:.4f} {significance})")
                            
                            elif is_numeric:
                                # 數值題分析
                                df_numeric = col_data_numeric.to_frame(name='value')
                                df_numeric['respondent_type'] = df_to_analyze.loc[df_numeric.index, 'respondent_type'].fillna('未知')
                                
                                # 繪製盒狀圖
                                fig = go.Figure()
                                colors = {'公司方': '#1f77b4', '投資方': '#ff7f0e', '未知': '#999999'}
                                
                                for resp_type in df_numeric['respondent_type'].unique():
                                    data_subset = df_numeric[df_numeric['respondent_type'] == resp_type]['value']
                                    fig.add_trace(go.Box(
                                        y=data_subset,
                                        name=resp_type,
                                        marker_color=colors.get(resp_type, '#cccccc'),
                                        boxmean='sd'
                                    ))
                                
                                fig.update_layout(
                                    title='數值分佈比較',
                                    yaxis_title='數值',
                                    template='plotly_white',
                                    height=400
                                )
                                st.plotly_chart(fig, use_container_width=True)
                                
                                # 統計摘要表
                                summary = df_numeric.groupby('respondent_type')['value'].describe()
                                st.dataframe(summary.style.format("{:.2f}"), use_container_width=True)
                                
                                # Mann-Whitney U 檢定
                                if 'p' in rec_info['統計結果']:
                                    p_val = rec_info['統計結果']['p']
                                    median_diff = rec_info['統計結果'].get('median_diff', 0)
                                    significance = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*"
                                    
                                    st.markdown("**統計檢定結果（Mann-Whitney U 檢定）：**")
                                    st.write(f"- p-value = {p_val:.4f} {significance}")
                                    st.write(f"- 中位數差異 = {median_diff:.2f}")
                                    
                                    if p_val < 0.05:
                                        st.success("✅ 公司方與投資方的數值分佈有顯著差異")
                                    else:
                                        st.info("ℹ️ 公司方與投資方的數值分佈無顯著差異")
                            
                            else:
                                # 類別題分析
                                s = col_data.astype(str)
                                s = s[~s.str.lower().str.contains('nan', na=False)]
                                
                                if not s.empty:
                                    df_cat = s.to_frame(name='category')
                                    df_cat['respondent_type'] = df_to_analyze.loc[df_cat.index, 'respondent_type'].fillna('未知')
                                    
                                    # 計算比例
                                    crosstab = pd.crosstab(df_cat['category'], df_cat['respondent_type'], normalize='columns') * 100
                                    
                                    # 智慧排序 x 軸
                                    sorted_index = smart_sort_categories(crosstab.index)
                                    crosstab = crosstab.reindex(sorted_index)
                                    
                                    if crosstab.shape[1] >= 2:
                                        # 繪製分組長條圖
                                        fig = go.Figure()
                                        colors = {'公司方': '#1f77b4', '投資方': '#ff7f0e', '未知': '#999999'}
                                        
                                        for resp_type in crosstab.columns:
                                            fig.add_trace(go.Bar(
                                                name=resp_type,
                                                x=crosstab.index,
                                                y=crosstab[resp_type],
                                                marker_color=colors.get(resp_type, '#cccccc'),
                                                text=[f"{v:.1f}%" for v in crosstab[resp_type]],
                                                textposition='auto'
                                            ))
                                        
                                        fig.update_layout(
                                            barmode='group',
                                            title='各類別在不同身分的分佈比例',
                                            xaxis_title='類別',
                                            yaxis_title='比例 (%)',
                                            template='plotly_white',
                                            height=400,
                                            xaxis_tickangle=-45,
                                            xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                                        )
                                        st.plotly_chart(fig, use_container_width=True)
                                        
                                        # 統計檢定
                                        if 'p' in rec_info['統計結果']:
                                            p_val = rec_info['統計結果']['p']
                                            significance = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*"
                                            
                                            st.markdown("**統計檢定結果（卡方檢定/Fisher精確檢定）：**")
                                            st.write(f"- p-value = {p_val:.4f} {significance}")
                                            
                                            if p_val < 0.05:
                                                st.success("✅ 公司方與投資方的分佈有顯著差異")
                                            else:
                                                st.info("ℹ️ 公司方與投資方的分佈無顯著差異")
                        
                        # === 分析2: 階段比較 (一階段 vs 二階段 vs 三階段) ===
                        if PHASE_COLUMN_NAME in df_to_analyze.columns and df_to_analyze[PHASE_COLUMN_NAME].notna().any():
                            phase_nunique = df_to_analyze.loc[col_data.index, PHASE_COLUMN_NAME].nunique()
                            
                            if phase_nunique > 1:
                                st.markdown("---")
                                st.markdown("#### 🔢 階段比較分析（一階段 vs 二階段 vs 三階段）")
                                
                                if is_multiselect:
                                    # 複選題階段分析
                                    exploded = col_data.astype(str).str.split('\n').explode().str.strip()
                                    exploded = exploded[(exploded != '') & (exploded != 'nan') & exploded.notna()]
                                    
                                    if not exploded.empty:
                                        df_exp = exploded.to_frame(name='option')
                                        df_exp['phase'] = df_to_analyze.loc[df_exp.index, PHASE_COLUMN_NAME].fillna('未標註')
                                        
                                        # 計算各選項在不同階段的比例
                                        crosstab_phase = pd.crosstab(df_exp['option'], df_exp['phase'], normalize='columns') * 100
                                        
                                        # 智慧排序 x 軸
                                        sorted_index = smart_sort_categories(crosstab_phase.index)
                                        crosstab_phase = crosstab_phase.reindex(sorted_index)
                                        
                                        # 繪製堆疊長條圖
                                        fig = go.Figure()
                                        colors = ['#2ca02c', '#d62728', '#9467bd', '#8c564b']
                                        
                                        for idx, phase in enumerate(sorted(crosstab_phase.columns)):
                                            fig.add_trace(go.Bar(
                                                name=str(phase),
                                                x=crosstab_phase.index,
                                                y=crosstab_phase[phase],
                                                marker_color=colors[idx % len(colors)],
                                                text=[f"{v:.1f}%" for v in crosstab_phase[phase]],
                                                textposition='auto'
                                            ))
                                        
                                        fig.update_layout(
                                            barmode='group',
                                            title='各選項在不同階段的選擇比例',
                                            xaxis_title='選項',
                                            yaxis_title='比例 (%)',
                                            template='plotly_white',
                                            height=500,
                                            xaxis_tickangle=-45,
                                            xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                                        )
                                        st.plotly_chart(fig, use_container_width=True)
                                        
                                        # 卡方檢定（檢查各選項在階段間是否有差異）
                                        st.markdown("**統計檢定結果（卡方檢定）：**")
                                        significant_options = []
                                        
                                        for opt in df_exp['option'].unique()[:10]:
                                            if pd.isna(opt):
                                                continue
                                            pres = df_to_analyze[topic].astype(str).fillna('').apply(
                                                lambda s: opt in [x.strip() for x in s.split('\n') if x.strip()]
                                            )
                                            table = pd.crosstab(pres, df_to_analyze.loc[pres.index, PHASE_COLUMN_NAME])
                                            
                                            if table.size > 0 and table.values.sum() > 0 and table.shape[0] >= 2 and table.shape[1] >= 2:
                                                try:
                                                    chi2, p, dof, exp = chi2_contingency(table)
                                                    if np.nanmin(exp) > 1 and p < 0.05:
                                                        significance = "***" if p < 0.001 else "**" if p < 0.01 else "*"
                                                        significant_options.append((opt, p, significance))
                                                except:
                                                    pass
                                        
                                        if significant_options:
                                            for opt, p, sig in significant_options[:5]:
                                                st.write(f"- 選項「{opt}」：不同階段間有顯著差異 (p = {p:.4f} {sig})")
                                        else:
                                            st.info("ℹ️ 各選項在不同階段間無顯著差異")
                                
                                elif is_numeric:
                                    # 數值題階段分析
                                    df_numeric_phase = col_data_numeric.to_frame(name='value')
                                    df_numeric_phase['phase'] = df_to_analyze.loc[df_numeric_phase.index, PHASE_COLUMN_NAME].fillna('未標註')
                                    
                                    # 繪製盒狀圖
                                    fig = go.Figure()
                                    colors = ['#2ca02c', '#d62728', '#9467bd', '#8c564b']
                                    
                                    for idx, phase in enumerate(sorted(df_numeric_phase['phase'].unique())):
                                        data_subset = df_numeric_phase[df_numeric_phase['phase'] == phase]['value']
                                        fig.add_trace(go.Box(
                                            y=data_subset,
                                            name=str(phase),
                                            marker_color=colors[idx % len(colors)],
                                            boxmean='sd'
                                        ))
                                    
                                    fig.update_layout(
                                        title='不同階段的數值分佈比較',
                                        yaxis_title='數值',
                                        template='plotly_white',
                                        height=400
                                    )
                                    st.plotly_chart(fig, use_container_width=True)
                                    
                                    # 統計摘要表
                                    summary_phase = df_numeric_phase.groupby('phase')['value'].describe()
                                    st.dataframe(summary_phase.style.format("{:.2f}"), use_container_width=True)
                                    
                                    # Kruskal-Wallis 檢定
                                    phases = df_numeric_phase['phase'].unique()
                                    if len(phases) >= 2:
                                        groups = [df_numeric_phase[df_numeric_phase['phase'] == p]['value'].values for p in phases]
                                        groups = [g for g in groups if len(g) > 0]
                                        
                                        if len(groups) >= 2:
                                            try:
                                                if len(groups) == 2:
                                                    stat, p_val = mannwhitneyu(groups[0], groups[1], alternative='two-sided')
                                                    test_name = "Mann-Whitney U 檢定"
                                                else:
                                                    stat, p_val = kruskal(*groups)
                                                    test_name = "Kruskal-Wallis 檢定"
                                                
                                                significance = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*"
                                                
                                                st.markdown(f"**統計檢定結果（{test_name}）：**")
                                                st.write(f"- p-value = {p_val:.4f} {significance}")
                                                
                                                if p_val < 0.05:
                                                    st.success("✅ 不同階段的數值分佈有顯著差異")
                                                else:
                                                    st.info("ℹ️ 不同階段的數值分佈無顯著差異")
                                            except Exception as e:
                                                st.warning(f"無法進行統計檢定：{str(e)}")
                                
                                else:
                                    # 類別題階段分析
                                    s = col_data.astype(str)
                                    s = s[~s.str.lower().str.contains('nan', na=False)]
                                    
                                    if not s.empty:
                                        df_cat_phase = s.to_frame(name='category')
                                        df_cat_phase['phase'] = df_to_analyze.loc[df_cat_phase.index, PHASE_COLUMN_NAME].fillna('未標註')
                                        
                                        # 計算比例
                                        crosstab_phase = pd.crosstab(df_cat_phase['category'], df_cat_phase['phase'], normalize='columns') * 100
                                        
                                        # 智慧排序 x 軸
                                        sorted_index = smart_sort_categories(crosstab_phase.index)
                                        crosstab_phase = crosstab_phase.reindex(sorted_index)
                                        
                                        # 繪製分組長條圖
                                        fig = go.Figure()
                                        colors = ['#2ca02c', '#d62728', '#9467bd', '#8c564b']
                                        
                                        for idx, phase in enumerate(sorted(crosstab_phase.columns)):
                                            fig.add_trace(go.Bar(
                                                name=str(phase),
                                                x=crosstab_phase.index,
                                                y=crosstab_phase[phase],
                                                marker_color=colors[idx % len(colors)],
                                                text=[f"{v:.1f}%" for v in crosstab_phase[phase]],
                                                textposition='auto'
                                            ))
                                        
                                        fig.update_layout(
                                            barmode='group',
                                            title='各類別在不同階段的分佈比例',
                                            xaxis_title='類別',
                                            yaxis_title='比例 (%)',
                                            template='plotly_white',
                                            height=400,
                                            xaxis_tickangle=-45,
                                            xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                                        )
                                        st.plotly_chart(fig, use_container_width=True)
                                        
                                        # 卡方檢定
                                        try:
                                            count_table = pd.crosstab(df_cat_phase['category'], df_cat_phase['phase'])
                                            chi2, p_val, dof, exp = chi2_contingency(count_table)
                                            
                                            significance = "***" if p_val < 0.001 else "**" if p_val < 0.01 else "*"
                                            
                                            st.markdown("**統計檢定結果（卡方檢定）：**")
                                            st.write(f"- p-value = {p_val:.4f} {significance}")
                                            
                                            if p_val < 0.05:
                                                st.success("✅ 不同階段的分佈有顯著差異")
                                            else:
                                                st.info("ℹ️ 不同階段的分佈無顯著差異")
                                        except Exception as e:
                                            st.warning(f"無法進行統計檢定：{str(e)}")
                        
                        # === 圖表說故事 ===
                        st.markdown("---")
                        st.markdown("#### 💡 分析洞察")
                        
                        insights = []
                        
                        # 根據統計結果生成洞察
                        if '顯著選項' in rec_info['統計結果']:
                            sig_count = rec_info['統計結果'].get('顯著選項數', 0)
                            insights.append(f"📌 本題有 {sig_count} 個選項在公司方與投資方之間呈現顯著差異，顯示兩者對此議題的看法或實務做法存在明顯不同。")
                        
                        if 'p' in rec_info['統計結果']:
                            p_val = rec_info['統計結果']['p']
                            if p_val < 0.001:
                                insights.append("📌 統計檢定顯示極度顯著差異 (p < 0.001)，建議在報告中重點探討造成差異的原因。")
                            elif p_val < 0.01:
                                insights.append("📌 統計檢定顯示高度顯著差異 (p < 0.01)，值得進一步分析不同群體的特性。")
                            elif p_val < 0.05:
                                insights.append("📌 統計檢定顯示顯著差異 (p < 0.05)，可在報告中提及此發現。")
                        
                        if rec_info['缺失率'] == "0.0%":
                            insights.append("📌 本題資料完整度極高（無缺失值），分析結果可信度高。")
                        
                        if insights:
                            for insight in insights:
                                st.write(insight)
                        else:
                            st.info("ℹ️ 本題未發現顯著的統計差異，但仍可作為描述性統計使用。")
        else:
            st.info("💡 目前沒有高優先順序（≥ 2）的題目，建議降低篩選標準或檢查資料品質。")
    else:
        st.warning("未找到具有顯著差異的題目")
    
    # === 新增：專業報告生成 ===
    if recommendations:
        st.markdown("---")
        st.markdown("### 📄 專業分析報告生成")
        st.info("✨ 為國發基金量身打造的專業分析報告，採用「執行摘要 → 方法論 → 主要發現 → 結論與建議」結構")
        
        if st.button("📊 生成完整分析報告", type="primary"):
            with st.spinner("正在生成專業報告..."):
                # 生成報告內容
                report = generate_professional_report(df_to_analyze, recommendations, cols_to_analyze, analysis_mode)
                
                # 顯示報告
                st.markdown("---")
                st.markdown(report, unsafe_allow_html=True)
                
                # 提供下載選項
                st.markdown("---")
                st.download_button(
                    label="💾 下載報告（Markdown 格式）",
                    data=report,
                    file_name=f"未上市櫃公司治理問卷分析報告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                    mime="text/markdown"
                )

# --- 題目顯示區 ---
st.markdown("---")
st.markdown("### 📝 題目分析與視覺化")

expand_all = st.checkbox("一鍵展開/收合所有題目", value=False, key="expand_all_toggle")
st.markdown("---")

# 逐題顯示
for i, col_name in enumerate(cols_to_analyze):
    if col_name not in df_to_analyze.columns:
        continue
        
    col_data = df_to_analyze[col_name].dropna()
    if col_data.empty:
        continue
        
    with st.expander(f"題目 {i+1}：{col_name}", expanded=expand_all):
        # 顯示樣本數
        st.caption(f"有效樣本數：{len(col_data)}")
        
        # 判斷題型
        is_multiselect = False
        if col_data.dtype == 'object':
            non_empty_data = col_data[col_data.astype(str) != '']
            if not non_empty_data.empty and non_empty_data.str.contains('\n').any():
                is_multiselect = True
        
        if is_multiselect:
            # 複選題
            st.markdown("##### 📊 複選題選項次數分佈")
            exploded = col_data.astype(str).str.split('\n').explode().str.strip()
            exploded = exploded[(exploded != '') & (exploded != 'nan') & exploded.notna()]
            
            if not exploded.empty:
                total_counts = exploded.value_counts().reset_index()
                total_counts.columns = ['選項', '次數']
                st.dataframe(total_counts, use_container_width=True)
                
                # 視覺化：如果有階段欄位則按階段分色堆疊
                if PHASE_COLUMN_NAME in df_to_analyze.columns and df_to_analyze[PHASE_COLUMN_NAME].notna().any() and df_to_analyze[PHASE_COLUMN_NAME].nunique() > 1:
                    st.markdown("##### 📈 各階段分佈（堆疊長條圖）")
                    exploded_df = exploded.to_frame(name='option')
                    exploded_df['phase'] = df_to_analyze.loc[exploded_df.index, PHASE_COLUMN_NAME].fillna('未標註階段')
                    pivot = exploded_df.groupby(['option', 'phase']).size().unstack(fill_value=0)
                    
                    # 智慧排序 x 軸
                    sorted_index = smart_sort_categories(pivot.index)
                    pivot = pivot.reindex(sorted_index)
                    
                    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                    fig = go.Figure()
                    for j, phase in enumerate(pivot.columns):
                        fig.add_trace(go.Bar(
                            x=pivot.index,
                            y=pivot[phase],
                            name=str(phase),
                            marker_color=colors[j % len(colors)]
                        ))
                    fig.update_layout(
                        barmode='stack', 
                        xaxis_tickangle=-45, 
                        template="plotly_white", 
                        height=500,
                        xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                    )
                    st.plotly_chart(fig, use_container_width=True, key=f"multi_{i}_{col_name[:20]}")
                else:
                    st.markdown("##### 📈 長條圖")
                    # 智慧排序 x 軸
                    sorted_index = smart_sort_categories(total_counts['選項'])
                    total_counts_sorted = total_counts.set_index('選項').reindex(sorted_index).reset_index()
                    
                    fig = go.Figure(data=[go.Bar(x=total_counts_sorted['選項'], y=total_counts_sorted['次數'])])
                    fig.update_layout(
                        xaxis_tickangle=-45, 
                        template="plotly_white", 
                        height=500,
                        xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                    )
                    st.plotly_chart(fig, use_container_width=True, key=f"multi_{i}_{col_name[:20]}")
            
            # 統計分析 - 複選題
            perform_comprehensive_statistical_analysis(df_to_analyze, col_data, col_name, is_numeric=False, is_multiselect=True)
        else:
            # 單選或數值題
            is_numeric = pd.api.types.is_numeric_dtype(col_data)
            if not is_numeric:
                numeric_version = pd.to_numeric(col_data, errors='coerce')
                if (numeric_version.notna().sum() / len(col_data) > 0.7):
                    is_numeric = True
                    col_data = numeric_version.dropna()
            
            if is_numeric:
                # 數值題
                st.markdown("##### 📊 數值統計摘要")
                st.dataframe(col_data.describe().to_frame().T.style.format("{:,.2f}"), use_container_width=True)
                
                # 盒狀圖：如果有階段欄位則按階段分組顯示
                if PHASE_COLUMN_NAME in df_to_analyze.columns and df_to_analyze[PHASE_COLUMN_NAME].notna().any() and df_to_analyze[PHASE_COLUMN_NAME].nunique() > 1:
                    st.markdown("##### 📦 盒狀圖（各階段比較）")
                    df_numeric = col_data.to_frame(name='value')
                    df_numeric['phase'] = df_to_analyze.loc[df_numeric.index, PHASE_COLUMN_NAME].fillna('未標註階段')
                    
                    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                    fig = go.Figure()
                    for j, phase in enumerate(sorted(df_numeric['phase'].unique())):
                        phase_data = df_numeric[df_numeric['phase'] == phase]['value']
                        fig.add_trace(go.Box(
                            y=phase_data,
                            name=str(phase),
                            marker_color=colors[j % len(colors)]
                        ))
                    fig.update_layout(template="plotly_white", height=400, showlegend=True)
                    st.plotly_chart(fig, use_container_width=True, key=f"num_{i}_{col_name[:20]}")
                else:
                    st.markdown("##### 📦 盒狀圖")
                    fig = go.Figure(data=[go.Box(y=col_data, name=col_name[:50])])
                    fig.update_layout(template="plotly_white", height=400)
                    st.plotly_chart(fig, use_container_width=True, key=f"num_{i}_{col_name[:20]}")
                
                # 統計分析 - 數值題
                perform_comprehensive_statistical_analysis(df_to_analyze, col_data, col_name, is_numeric=True, is_multiselect=False)
            else:
                # 類別題
                st.markdown("##### 📊 類別次數分佈")
                s = col_data.astype(str)
                s = s[~s.str.lower().str.contains('nan', na=False)]
                
                if not s.empty:
                    total = s.value_counts().reset_index()
                    total.columns = ['選項', '次數']
                    st.dataframe(total, use_container_width=True)
                    
                    # 視覺化：如果有階段欄位則按階段分色堆疊
                    if PHASE_COLUMN_NAME in df_to_analyze.columns and df_to_analyze[PHASE_COLUMN_NAME].notna().any() and df_to_analyze[PHASE_COLUMN_NAME].nunique() > 1:
                        st.markdown("##### 📈 各階段分佈（堆疊長條圖）")
                        df_pair = s.to_frame(name='ans')
                        df_pair['phase'] = df_to_analyze.loc[df_pair.index, PHASE_COLUMN_NAME].fillna('未標註階段')
                        pivot = df_pair.groupby(['ans', 'phase']).size().unstack(fill_value=0)
                        
                        # 智慧排序 x 軸
                        sorted_index = smart_sort_categories(pivot.index)
                        pivot = pivot.reindex(sorted_index)
                        
                        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                        fig = go.Figure()
                        for j, phase in enumerate(pivot.columns):
                            fig.add_trace(go.Bar(
                                x=pivot.index,
                                y=pivot[phase],
                                name=str(phase),
                                marker_color=colors[j % len(colors)]
                            ))
                        fig.update_layout(
                            barmode='stack', 
                            xaxis_tickangle=-45, 
                            template="plotly_white", 
                            height=500,
                            xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                        )
                        st.plotly_chart(fig, use_container_width=True, key=f"cat_{i}_{col_name[:20]}")
                    else:
                        st.markdown("##### 📈 長條圖")
                        # 智慧排序 x 軸
                        sorted_index = smart_sort_categories(total['選項'])
                        total_sorted = total.set_index('選項').reindex(sorted_index).reset_index()
                        
                        fig = go.Figure(data=[go.Bar(x=total_sorted['選項'], y=total_sorted['次數'])])
                        fig.update_layout(
                            xaxis_tickangle=-45, 
                            template="plotly_white", 
                            height=500,
                            xaxis={'categoryorder': 'array', 'categoryarray': sorted_index}
                        )
                        st.plotly_chart(fig, use_container_width=True, key=f"cat_{i}_{col_name[:20]}")
                
                # 統計分析 - 類別題
                perform_comprehensive_statistical_analysis(df_to_analyze, col_data, col_name, is_numeric=False, is_multiselect=False)
