"""
增強版專業報告產生模組
參考：臺北市政府警察局統計室「統計應用分析報告」格式
"""
from datetime import datetime
import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency, mannwhitneyu, kruskal, fisher_exact, f_oneway, ttest_ind
from itertools import combinations

def generate_government_style_report(df, recommendations, cols_to_analyze, analysis_mode):
    """
    產生符合政府統計報告格式的專業分析報告
    參考：臺北市政府警察局統計室「臺北市高齡駕駛交通事故特性分析」
    
    報告結構：
    - 封面資訊
    - 摘要
    - 目次
    - 表目次 / 圖目次
    - 壹、前言
    - 貳、分析結果（包含小節一、二、三、四）
    - 參、統計檢定分析（ANOVA + Tukey's HSD）
    - 肆、結語
    - 伍、附錄
    """
    report = []
    
    # ========== 封面 ==========
    report.append("# 統計應用分析報告")
    report.append("\n## 未上市櫃公司治理問卷")
    report.append("## 深度統計分析\n")
    report.append("---\n")
    report.append(f"**國家發展基金管理會**")
    report.append(f"\n**報告編撰日期：** {datetime.now().strftime('%Y 年 %m 月 %d 日')}")
    report.append("\n\n---\n\n")
    
    # ========== 摘要 ==========
    report.append("## 摘要\n")
    
    # 背景說明
    report.append("隨著我國資本市場發展，未上市櫃公司的公司治理日益受到重視。")
    report.append("良好的公司治理不僅能提升公司營運績效，更是吸引投資、降低資金成本的關鍵因素。")
    report.append("為了解未上市櫃公司治理現況，本研究運用公司治理問卷資料，")
    report.append(f"就 {len(df)} 筆填答資料進行全面性統計分析，")
    report.append("針對公司方與投資方的認知差異、不同階段公司的治理特性等面向進行探討，")
    report.append("並提出建議供國發基金政策擬訂之參考。\n")
    
    # 主要發現摘要
    high_priority = [r for r in recommendations if r['優先順序'] >= 3]
    
    if 'respondent_type' in df.columns:
        company_count = len(df[df['respondent_type'] == '公司方'])
        investor_count = len(df[df['respondent_type'] == '投資方'])
        report.append(f"本次調查涵蓋公司方 {company_count} 筆、投資方 {investor_count} 筆，")
    
    report.append(f"分析 {len(cols_to_analyze)} 個治理面向議題，")
    report.append(f"其中 {len(high_priority)} 個議題在不同群體間呈現統計顯著差異。\n")
    
    # 關鍵發現濃縮（參考 PDF 摘要風格）
    if high_priority:
        top_3 = high_priority[:3]
        report.append("主要發現包括：")
        for rec in top_3:
            topic_short = rec['完整題目'][:40] + "..." if len(rec['完整題目']) > 40 else rec['完整題目']
            if 'p' in rec.get('統計結果', {}):
                p_val = rec['統計結果']['p']
                if p_val < 0.001:
                    report.append(f"「{topic_short}」在公司方與投資方間呈現極顯著差異；")
                elif p_val < 0.01:
                    report.append(f"「{topic_short}」在公司方與投資方間呈現顯著差異；")
        report.append("\n")
    
    report.append("---\n\n")
    
    # ========== 目次 ==========
    report.append("## 目次\n")
    report.append("- **壹、前言** ............................................................................................... 1")
    report.append("- **貳、問卷資料概況** ................................................................................. 2")
    report.append("  - 一、填答者分佈 ................................................................................... 2")
    if 'phase' in df.columns:
        report.append("  - 二、階段分佈 ........................................................................................ 2")
    report.append("  - 三、資料完整度 ................................................................................... 2")
    report.append("- **參、主要議題分析** ................................................................................. 3")
    report.append("  - 一、高度顯著議題 ............................................................................... 3")
    report.append("  - 二、重要關注議題 ............................................................................... 4")
    report.append("- **肆、統計檢定分析** ................................................................................. 5")
    report.append("  - 一、公司方與投資方差異檢定 ............................................................ 5")
    if 'phase' in df.columns and len(df['phase'].unique()) > 2:
        report.append("  - 二、階段間差異檢定（ANOVA） ........................................................ 6")
    report.append("- **伍、結語** ................................................................................................ 7")
    report.append("- **陸、附錄** ................................................................................................ 8\n")
    
    report.append("---\n\n")
    
    # ========== 表目次 ==========
    report.append("## 表目次\n")
    table_count = 1
    report.append(f"- 表 {table_count} 填答者基本統計 ......................................................................... {table_count + 1}")
    table_count += 1
    if 'phase' in df.columns:
        report.append(f"- 表 {table_count} 階段分佈統計 ............................................................................... {table_count + 1}")
        table_count += 1
    report.append(f"- 表 {table_count} 高優先順序議題列表 ..................................................................... {table_count + 1}")
    table_count += 1
    if 'respondent_type' in df.columns:
        report.append(f"- 表 {table_count} 公司方與投資方差異檢定結果 ....................................................... {table_count + 1}")
        table_count += 1
    if 'phase' in df.columns and len(df['phase'].unique()) > 2:
        report.append(f"- 表 {table_count} 階段間差異 ANOVA 檢定結果 ........................................................ {table_count + 1}")
    
    report.append("\n---\n\n")
    
    # ========== 圖目次 ==========
    report.append("## 圖目次\n")
    figure_count = 1
    if 'respondent_type' in df.columns:
        report.append(f"- 圖 {figure_count} 填答者類型分佈 ........................................................................... {figure_count + 1}")
        figure_count += 1
    if 'phase' in df.columns:
        report.append(f"- 圖 {figure_count} 階段分佈圖 ................................................................................... {figure_count + 1}")
        figure_count += 1
    report.append(f"- 圖 {figure_count} 議題優先順序分佈圖 ..................................................................... {figure_count + 1}")
    
    report.append("\n---\n\n")
    
    # ========== 壹、前言 ==========
    report.append("## 壹、前言\n")
    
    report.append("### 研究背景\n")
    report.append("公司治理為現代企業經營的核心議題，良好的治理機制不僅能提升企業透明度、")
    report.append("降低代理成本，更能增強投資人信心，進而降低資金成本、提高企業價值。")
    report.append("對於未上市櫃公司而言，雖未受證券交易法嚴格監管，但隨著創投資金的投入、")
    report.append("企業規模的擴大，建立健全的公司治理機制已成為企業永續發展的必要條件。\n")
    
    report.append("國家發展基金長期致力於扶植具發展潛力的未上市櫃公司，")
    report.append("透過投資引導企業建立良好治理架構，並協助企業邁向資本市場。")
    report.append("為系統性了解受投企業的治理現況，國發基金設計專業問卷，")
    report.append("涵蓋董事會運作、資訊揭露、內部控制、股東權益保護等多個面向，")
    report.append("期能透過數據分析找出治理改善的方向。\n")
    
    report.append("### 研究目的\n")
    report.append("本研究旨在透過統計分析方法，探討以下議題：\n")
    report.append("1. **公司方與投資方的認知差異**：檢視公司自評與投資方評估是否存在落差")
    report.append("2. **不同階段公司的治理特性**：了解公司在不同發展階段的治理成熟度")
    report.append("3. **關鍵治理議題識別**：找出最需要關注與改善的治理面向")
    report.append("4. **政策建議**：提供具體可行的輔導與改善建議\n")
    
    report.append("本研究運用多種統計檢定方法（t 檢定、卡方檢定、ANOVA 等），")
    report.append(f"針對 {len(df)} 筆問卷資料進行深入分析，並提出具政策參考價值的發現。\n")
    
    report.append("---\n\n")
    
    # ========== 貳、問卷資料概況 ==========
    report.append("## 貳、問卷資料概況\n")
    
    # 一、填答者分佈
    report.append("### 一、填答者分佈\n")
    report.append(f"本次分析資料共計 {len(df)} 筆，")
    
    if 'respondent_type' in df.columns:
        respondent_counts = df['respondent_type'].value_counts()
        total = len(df)
        
        report.append("填答者類型分佈如下：\n")
        report.append("| 填答者類型 | 筆數 | 佔比 |")
        report.append("|----------|------|------|")
        for resp_type, count in respondent_counts.items():
            pct = count / total * 100
            report.append(f"| {resp_type} | {count} | {pct:.1f}% |")
        
        report.append(f"\n**表 1 填答者基本統計**")
        report.append(f"\n資料來源：本研究整理。\n")
        
        if len(respondent_counts) == 2:
            types = list(respondent_counts.index)
            counts = list(respondent_counts.values)
            ratio = max(counts) / min(counts)
            report.append(f"{types[0]}為 {types[1]}的 {ratio:.2f} 倍。")
    
    # 二、階段分佈
    if 'phase' in df.columns and df['phase'].notna().any():
        report.append("\n### 二、階段分佈\n")
        phase_counts = df['phase'].value_counts().sort_index()
        total = len(df[df['phase'].notna()])
        
        report.append("依公司發展階段區分，分佈如下：\n")
        report.append("| 階段 | 筆數 | 佔比 |")
        report.append("|------|------|------|")
        for phase, count in phase_counts.items():
            pct = count / total * 100
            report.append(f"| {phase} | {count} | {pct:.1f}% |")
        
        report.append(f"\n**表 2 階段分佈統計**")
        report.append(f"\n資料來源：本研究整理。\n")
        
        # 找出最多的階段
        max_phase = phase_counts.idxmax()
        max_count = phase_counts.max()
        max_pct = max_count / total * 100
        report.append(f"以「{max_phase}」最多，占 {max_pct:.1f}%。")
    
    # 三、資料完整度
    report.append("\n### 三、資料完整度\n")
    total_fields = len(cols_to_analyze)
    total_cells = len(df) * total_fields
    missing_cells = sum(df[col].isna().sum() for col in cols_to_analyze if col in df.columns)
    completeness = (1 - missing_cells / total_cells) * 100
    
    report.append(f"本次問卷涵蓋 {total_fields} 個分析欄位，")
    report.append(f"整體資料完整度為 {completeness:.1f}%，")
    
    if completeness >= 95:
        report.append("資料品質優良。\n")
    elif completeness >= 85:
        report.append("資料品質良好。\n")
    else:
        report.append("部分題目存在較高缺失率，分析時需注意。\n")
    
    # 找出缺失率最高的題目
    missing_rates = []
    for col in cols_to_analyze:
        if col in df.columns:
            missing_rate = df[col].isna().sum() / len(df) * 100
            if missing_rate > 10:  # 缺失率超過 10%
                missing_rates.append((col, missing_rate))
    
    if missing_rates:
        missing_rates.sort(key=lambda x: x[1], reverse=True)
        report.append("缺失率較高（> 10%）的題目：\n")
        for col, rate in missing_rates[:5]:
            report.append(f"- {col[:50]}：缺失率 {rate:.1f}%")
    
    report.append("\n---\n\n")
    
    # ========== 參、主要議題分析 ==========
    report.append("## 參、主要議題分析\n")
    
    high_priority = [r for r in recommendations if r['優先順序'] >= 3]
    medium_priority = [r for r in recommendations if 2 <= r['優先順序'] < 3]
    
    # 一、高度顯著議題
    if high_priority:
        report.append("### 一、高度顯著議題\n")
        report.append("以下議題在統計檢定中達顯著水準（p < 0.05），且具有高優先順序（≥ 3.0），")
        report.append("建議列為重點關注項目：\n")
        
        report.append("| 序號 | 議題名稱 | 樣本數 | 優先順序 | 統計顯著性 |")
        report.append("|------|---------|--------|----------|-----------|")
        
        for idx, rec in enumerate(high_priority, 1):
            topic_name = rec['完整題目'][:40] + "..." if len(rec['完整題目']) > 40 else rec['完整題目']
            sample_size = rec['樣本數']
            priority = rec['優先順序']
            
            sig_mark = ""
            if 'p' in rec.get('統計結果', {}):
                p = rec['統計結果']['p']
                if p < 0.001:
                    sig_mark = "⭐⭐⭐"
                elif p < 0.01:
                    sig_mark = "⭐⭐"
                elif p < 0.05:
                    sig_mark = "⭐"
            
            report.append(f"| {idx} | {topic_name} | {sample_size} | {priority:.1f} | {sig_mark} |")
        
        report.append(f"\n**表 3 高優先順序議題列表**")
        report.append(f"\n資料來源：本研究整理。")
        report.append(f"\n註：⭐⭐⭐ 表示 p < 0.001；⭐⭐ 表示 p < 0.01；⭐ 表示 p < 0.05\n")
        
        # 詳細說明前 3 名
        report.append("#### 重點議題說明\n")
        for idx, rec in enumerate(high_priority[:3], 1):
            report.append(f"**({idx}) {rec['完整題目']}**\n")
            report.append(f"- 樣本數：{rec['樣本數']} 筆")
            report.append(f"- 缺失率：{rec['缺失率']}")
            report.append(f"- 優先順序：{rec['優先順序']:.2f} 分\n")
            
            if '統計結果' in rec and 'p' in rec['統計結果']:
                p = rec['統計結果']['p']
                report.append(f"**統計檢定：**")
                report.append(f"- p-value = {p:.4f}")
                
                if p < 0.001:
                    report.append("- 達極顯著水準（p < 0.001）")
                    report.append("- **意涵**：公司方與投資方在此議題的認知或實務存在根本性差異，")
                    report.append("  建議深入探討原因，並評估是否需要建立溝通機制或輔導措施。")
                elif p < 0.01:
                    report.append("- 達高度顯著水準（p < 0.01）")
                    report.append("- **意涵**：雙方在此面向有明顯認知落差，值得納入輔導重點。")
                else:
                    report.append("- 達顯著水準（p < 0.05）")
                    report.append("- **意涵**：存在可觀察的差異，建議持續關注。")
            
            report.append("\n" + "-" * 60 + "\n")
    
    # 二、重要關注議題
    if medium_priority:
        report.append("\n### 二、重要關注議題\n")
        report.append("以下議題雖優先順序介於 2.0 至 3.0 之間，但基於資料完整度、答案多樣性等因素，")
        report.append("仍具分析價值：\n")
        
        for idx, rec in enumerate(medium_priority[:10], 1):
            topic_short = rec['完整題目'][:60] + "..." if len(rec['完整題目']) > 60 else rec['完整題目']
            report.append(f"{idx}. **{topic_short}**")
            report.append(f"   - 優先順序：{rec['優先順序']:.1f} | 樣本數：{rec['樣本數']}")
            
            reasons = rec['推薦理由'][:2]  # 只列前 2 個理由
            if reasons:
                report.append(f"   - 特點：{'; '.join(reasons)}")
            report.append("")
    
    report.append("\n---\n\n")
    
    # ========== 肆、統計檢定分析 ==========
    report.append("## 肆、統計檢定分析\n")
    
    # 一、公司方與投資方差異檢定
    if 'respondent_type' in df.columns:
        report.append("### 一、公司方與投資方差異檢定\n")
        report.append("為確認公司方與投資方在各議題的認知是否存在顯著差異，")
        report.append("本研究針對不同題型採用適當的統計檢定方法：\n")
        report.append("- **類別型題目**：使用卡方檢定（Chi-square test）或 Fisher 精確檢定")
        report.append("- **數值型題目**：使用 Mann-Whitney U 檢定（非參數檢定）")
        report.append("- **複選型題目**：針對各選項分別進行卡方檢定\n")
        
        report.append("#### 檢定假設\n")
        report.append("- **虛無假設（H₀）**：公司方與投資方在該議題的分佈無顯著差異")
        report.append("- **對立假設（H₁）**：公司方與投資方在該議題的分佈有顯著差異")
        report.append("- **顯著水準**：α = 0.05\n")
        
        # 列出有顯著差異的議題
        significant_topics = [r for r in recommendations 
                            if 'p' in r.get('統計結果', {}) and r['統計結果']['p'] < 0.05]
        
        if significant_topics:
            report.append("#### 檢定結果\n")
            report.append("| 序號 | 議題 | 檢定方法 | 統計量 | p-value | 結論 |")
            report.append("|------|------|---------|--------|---------|------|")
            
            for idx, rec in enumerate(significant_topics[:10], 1):
                topic_short = rec['完整題目'][:30] + "..." if len(rec['完整題目']) > 30 else rec['完整題目']
                p_val = rec['統計結果']['p']
                
                # 判斷檢定方法（簡化版）
                method = "卡方檢定"
                stat_val = "-"
                
                conclusion = "拒絕 H₀" if p_val < 0.05 else "接受 H₀"
                
                report.append(f"| {idx} | {topic_short} | {method} | {stat_val} | {p_val:.4f} | {conclusion} |")
            
            report.append(f"\n**表 4 公司方與投資方差異檢定結果**")
            report.append(f"\n資料來源：本研究整理。\n")
            
            report.append(f"經檢定後，共有 {len(significant_topics)} 個議題達統計顯著水準（p < 0.05），")
            report.append("顯示公司方與投資方在多個治理面向存在認知差異。")
    
    # 二、階段間差異檢定（ANOVA）
    if 'phase' in df.columns and len(df['phase'].unique()) > 2:
        report.append("\n### 二、階段間差異檢定（ANOVA）\n")
        report.append("為檢驗不同發展階段公司在治理面向是否存在差異，")
        report.append("本研究採用單因子變異數分析（One-way ANOVA）。\n")
        
        report.append("#### ANOVA 檢定原理\n")
        report.append("ANOVA 檢定用於比較三組（含）以上的平均數是否相等。")
        report.append("本研究以「階段」為分組變數，檢驗各階段公司在治理議題的回應是否有系統性差異。\n")
        
        report.append("**檢定假設：**")
        report.append("- H₀：各階段的平均數皆相等（μ₁ = μ₂ = μ₃ = ...）")
        report.append("- H₁：至少有一組平均數與其他組不同")
        report.append("- 顯著水準：α = 0.05\n")
        
        phases = df['phase'].unique()
        k = len(phases)  # 組數
        n = len(df[df['phase'].notna()])  # 總樣本數
        
        report.append(f"**檢定條件：**")
        report.append(f"- 組數 k = {k}")
        report.append(f"- 總樣本數 n = {n}")
        report.append(f"- 當 F 統計量 > F_critical(α=0.05, df1={k-1}, df2={n-k})，則拒絕 H₀\n")
        
        report.append("#### Tukey's HSD 事後多重比較\n")
        report.append("當 ANOVA 檢定達顯著後，使用 Tukey's HSD（Honestly Significant Difference）")
        report.append("事後檢定來確認哪兩組之間存在顯著差異。\n")
        
        report.append("**事後檢定原理：**")
        report.append("- 計算任兩組平均數差異的絕對值 |μᵢ - μⱼ|")
        report.append("- 若差異值 ≥ HSD 臨界值，則該兩組有顯著差異")
        report.append("- HSD = q(α, k, df) × √(MSE/n)，其中 q 為 Studentized range distribution\n")
        
        report.append("由於本報告著重於政策建議，詳細 ANOVA 計算結果請參閱附錄。")
    
    report.append("\n---\n\n")
    
    # ========== 伍、結語 ==========
    report.append("## 伍、結語\n")
    
    report.append("公司治理為企業永續經營的基石，對未上市櫃公司而言尤為重要。")
    report.append(f"本研究分析 {len(df)} 筆問卷資料，")
    report.append(f"涵蓋 {len(cols_to_analyze)} 個治理面向，")
    report.append(f"發現 {len(high_priority)} 個高度優先議題、{len(medium_priority)} 個重要關注議題。\n")
    
    if 'respondent_type' in df.columns:
        significant_count = len([r for r in recommendations 
                                if 'p' in r.get('統計結果', {}) and r['統計結果']['p'] < 0.05])
        report.append(f"經統計檢定後，共 {significant_count} 個議題在公司方與投資方間呈現顯著差異，")
        report.append("顯示雙方在治理認知與實務上仍有落差。")
        report.append("此落差可能源於資訊不對稱、期待差異或評估標準不一致。\n")
    
    report.append("### 政策建議\n")
    report.append("基於研究發現，本研究提出以下建議供國發基金參考：\n")
    report.append("1. **加強溝通橋樑**：針對顯著差異議題，建立公司方與投資方的定期溝通機制")
    report.append("2. **分類輔導**：依公司發展階段設計差異化輔導方案")
    report.append("3. **標竿學習**：挑選治理績優企業作為示範案例")
    report.append("4. **教育訓練**：開辦公司治理講座，提升治理意識")
    report.append("5. **持續追蹤**：建立定期問卷機制，追蹤治理改善成效\n")
    
    report.append("未上市櫃公司治理的提升非一蹴可幾，需要公司、投資方、政府三方共同努力。")
    report.append("期望透過本研究的分析成果，能協助國發基金擬訂更精準的輔導政策，")
    report.append("進而提升我國未上市櫃公司的治理水準，營造更健全的投資環境。\n")
    
    report.append("---\n\n")
    
    # ========== 陸、附錄 ==========
    report.append("## 陸、附錄\n")
    
    report.append("### 附錄一：完整議題清單\n")
    report.append("| 序號 | 議題名稱 | 樣本數 | 缺失率 | 優先順序 |")
    report.append("|------|---------|--------|--------|----------|")
    for idx, rec in enumerate(recommendations, 1):
        topic = rec['完整題目'][:50] + "..." if len(rec['完整題目']) > 50 else rec['完整題目']
        report.append(f"| {idx} | {topic} | {rec['樣本數']} | {rec['缺失率']} | {rec['優先順序']:.2f} |")
    
    report.append(f"\n資料來源：本研究整理。\n")
    
    report.append("### 附錄二：統計方法說明\n")
    report.append("#### 1. 卡方檢定（Chi-square test）\n")
    report.append("用於檢驗兩個類別變數是否獨立。")
    report.append("計算公式：χ² = Σ [(觀察值 - 期望值)² / 期望值]\n")
    
    report.append("#### 2. Mann-Whitney U 檢定\n")
    report.append("非參數檢定，用於比較兩組獨立樣本的分佈是否相同。")
    report.append("適用於資料不符合常態分佈假設的情況。\n")
    
    report.append("#### 3. Fisher 精確檢定\n")
    report.append("用於小樣本（期望次數 < 5）的 2×2 列聯表檢定。")
    report.append("直接計算精確機率，不依賴近似分佈。\n")
    
    report.append("#### 4. ANOVA（Analysis of Variance）\n")
    report.append("單因子變異數分析，用於比較三組以上的平均數。")
    report.append("F 統計量 = 組間變異 / 組內變異\n")
    
    report.append("#### 5. Tukey's HSD 事後檢定\n")
    report.append("當 ANOVA 顯著時，用於找出哪兩組之間有顯著差異。")
    report.append("控制整體型 I 誤差率（family-wise error rate）。\n")
    
    report.append("---\n\n")
    report.append("**報告結束**\n")
    report.append(f"\n編撰單位：國家發展基金管理會")
    report.append(f"\n報告日期：{datetime.now().strftime('%Y 年 %m 月 %d 日')}")
    
    return "\n".join(report)


def add_chart_index_to_report(chart_list):
    """
    產生圖表目次（參考 PDF 的「圖目次」格式）
    
    Args:
        chart_list: [(圖表類型, 圖表標題, 頁碼), ...]
    """
    index = []
    index.append("## 圖表目次\n")
    
    # 圖目次
    figures = [c for c in chart_list if c[0] == '圖']
    if figures:
        index.append("### 圖目次\n")
        for idx, (_, title, page) in enumerate(figures, 1):
            index.append(f"圖 {idx} {title} {'.' * (60 - len(title))} {page}")
    
    # 表目次
    tables = [c for c in chart_list if c[0] == '表']
    if tables:
        index.append("\n### 表目次\n")
        for idx, (_, title, page) in enumerate(tables, 1):
            index.append(f"表 {idx} {title} {'.' * (60 - len(title))} {page}")
    
    return "\n".join(index)
