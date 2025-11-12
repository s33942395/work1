"""
描述性統計報告生成器 - 配合 docx 格式
輸出為 Word 格式，包含圖表、表格、統計檢定
"""
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency, kruskal, mannwhitneyu, fisher_exact
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
from datetime import datetime
import os
import re

def smart_sort_categories(categories):
    """
    智慧排序類別資料 - 與 cloud_app.py 完全一致
    處理百分比、數值範圍、階段、是否等特殊格式
    """
    if len(categories) == 0:
        return []
    
    categories_list = list(categories)
    
    def sort_key(item):
        item_str = str(item).strip()
        
        # 1. 百分比範圍
        percent_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*[%％]', item_str)
        if percent_match:
            return (0, float(percent_match.group(1)))
        
        single_percent = re.match(r'(\d+\.?\d*)\s*[%％]', item_str)
        if single_percent:
            return (0, float(single_percent.group(1)))
        
        # 2. 年份範圍
        year_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*年', item_str)
        if year_match:
            return (1, float(year_match.group(1)))
        
        # 3. 金額範圍
        money_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*[萬億]', item_str)
        if money_match:
            return (2, float(money_match.group(1)))
        
        # 4. 月份範圍
        month_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*個?月', item_str)
        if month_match:
            return (3, float(month_match.group(1)))
        
        # 5. 人數範圍
        people_match = re.match(r'(\d+\.?\d*)\s*[-~到至]\s*(\d+\.?\d*)\s*人', item_str)
        if people_match:
            return (4, float(people_match.group(1)))
        
        # 6. 頻率
        freq_order = {'每週': 1, '每月': 2, '每季': 3, '每半年': 4, '每年': 5, '不定期': 6, '無': 7}
        for key, value in freq_order.items():
            if key in item_str:
                return (5, value)
        
        # 7. 階段
        if '第一階段' in item_str or '階段1' in item_str:
            return (6, 1)
        if '第二階段' in item_str or '階段2' in item_str:
            return (6, 2)
        if '第三階段' in item_str or '階段3' in item_str:
            return (6, 3)
        
        # 8. 是/否
        if item_str in ['是', 'Yes', 'yes', '有']:
            return (7, 1)
        if item_str in ['否', 'No', 'no', '無']:
            return (7, 2)
        
        # 9. 一般文字
        return (10, item_str)
    
    try:
        sorted_list = sorted(categories_list, key=sort_key)
        return sorted_list
    except:
        return categories_list

def add_heading_with_style(doc, text, level=1):
    """新增標題並設定樣式"""
    heading = doc.add_heading(text, level=level)
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return heading

def add_statistics_table(doc, data_dict, title=""):
    """
    新增政府統計風格的完整表格
    包含標題、資料來源、製表單位等資訊
    """
    if title:
        # 表格標題（置中）
        title_para = doc.add_paragraph(title)
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.runs[0].font.bold = True
        title_para.runs[0].font.size = Pt(12)
    
    # 創建表格
    table = doc.add_table(rows=1, cols=len(data_dict['columns']))
    table.style = 'Light Grid Accent 1'
    
    # 標題列（加粗、置中）
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(data_dict['columns']):
        hdr_cells[i].text = str(col_name)
        # 設定標題列樣式
        for paragraph in hdr_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(11)
    
    # 資料列
    for row_data in data_dict['data']:
        row_cells = table.add_row().cells
        for i, cell_value in enumerate(row_data):
            row_cells[i].text = str(cell_value)
            # 數值靠右對齊，文字靠左對齊
            if i > 0:  # 第一欄（類別）靠左，其他欄靠右
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # 表格備註（政府統計風格）
    doc.add_paragraph()
    note_para = doc.add_paragraph()
    note_para.add_run('資料來源：').font.size = Pt(9)
    note_para.add_run('問卷調查資料').font.size = Pt(9)
    note_para.runs[0].font.bold = True
    
    doc.add_paragraph()  # 空行
    return table

def save_plotly_as_image(fig, filename):
    """儲存 Plotly 圖表為圖片（PNG格式）"""
    try:
        # 嘗試儲存圖表，設定較長的超時時間
        fig.write_image(filename, width=1000, height=600, scale=2, engine="kaleido")
        return True
    except Exception as e:
        print(f"圖表儲存失敗: {e}")
        # 嘗試使用較低品質設定
        try:
            print("嘗試使用較低品質設定...")
            fig.write_image(filename, width=800, height=500, scale=1, engine="kaleido")
            return True
        except Exception as e2:
            print(f"圖表儲存再次失敗: {e2}")
            return False

def create_bar_chart(crosstab, crosstab_pct, title, categories):
    """
    創建長條圖（公司方 vs 投資方比較）- 與 cloud_app.py 完全一致
    使用智慧排序、plotly_white 模板、正確配色
    """
    # 智慧排序 X 軸
    sorted_categories = smart_sort_categories(categories)
    
    # 重新排序數據
    crosstab_sorted = crosstab.reindex(sorted_categories)
    crosstab_pct_sorted = crosstab_pct.reindex(sorted_categories)
    
    fig = go.Figure()
    
    # 使用與 app 一致的配色方案
    colors = {'公司方': '#1f77b4', '投資方': '#ff7f0e', '未知': '#999999'}
    
    # 公司方數據 - 使用百分比作為 Y 軸
    if '公司方' in crosstab_pct_sorted.columns:
        company_pct = crosstab_pct_sorted['公司方'].values
        
        fig.add_trace(go.Bar(
            name='公司方',
            x=sorted_categories,
            y=company_pct,
            marker_color=colors['公司方'],
            text=[f"{p:.1f}%" for p in company_pct],
            textposition='auto'
        ))
    
    # 投資方數據 - 使用百分比作為 Y 軸
    if '投資方' in crosstab_pct_sorted.columns:
        investor_pct = crosstab_pct_sorted['投資方'].values
        
        fig.add_trace(go.Bar(
            name='投資方',
            x=sorted_categories,
            y=investor_pct,
            marker_color=colors['投資方'],
            text=[f"{p:.1f}%" for p in investor_pct],
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
        xaxis={'categoryorder': 'array', 'categoryarray': sorted_categories},
        font=dict(family='Microsoft JhengHei', size=12)
    )
    
    return fig

def create_phase_chart(phase_crosstab, phase_crosstab_pct, title, categories, phases):
    """
    創建階段比較長條圖
    """
    fig = go.Figure()
    
    colors = ['rgb(55, 83, 109)', 'rgb(26, 118, 255)', 'rgb(50, 171, 96)']
    
    for i, phase in enumerate(phases):
        if phase in phase_crosstab.columns:
            values = [phase_crosstab.loc[cat, phase] if cat in phase_crosstab.index else 0 for cat in categories]
            pct = [phase_crosstab_pct.loc[cat, phase] if cat in phase_crosstab_pct.index else 0 for cat in categories]
            
            fig.add_trace(go.Bar(
                name=phase,
                x=categories,
                y=values,
                text=[f'{v}<br>({p:.1f}%)' for v, p in zip(values, pct)],
                textposition='auto',
                marker_color=colors[i % len(colors)]
            ))
    
    fig.update_layout(
        title=dict(text=title, font=dict(size=16, family='Microsoft JhengHei')),
        xaxis=dict(title='', tickfont=dict(size=11, family='Microsoft JhengHei')),
        yaxis=dict(title='人數', tickfont=dict(size=11)),
        barmode='group',
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Microsoft JhengHei'),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    
    return fig

def calculate_chi_square(df, col_name, group_col='respondent_type'):
    """計算卡方檢定"""
    try:
        crosstab = pd.crosstab(df[col_name], df[group_col])
        if crosstab.size > 0 and crosstab.values.sum() > 0:
            if crosstab.shape == (2, 2) and crosstab.values.sum() < 20:
                # 小樣本使用 Fisher 精確檢定
                oddsratio, p = fisher_exact(crosstab)
                return {'method': 'Fisher精確檢定', 'statistic': oddsratio, 'p_value': p}
            else:
                chi2, p, dof, expected = chi2_contingency(crosstab)
                return {'method': '卡方檢定', 'statistic': chi2, 'p_value': p, 'dof': dof}
    except Exception as e:
        print(f"統計檢定失敗: {e}")
    return None

def generate_descriptive_report_word(df, output_filename="問卷描述性統計報告_改進版.docx"):
    """
    生成描述性統計 Word 報告
    參考原始 docx 格式，加入圖表、表格、統計檢定
    """
    
    # 創建 Word 文件
    doc = Document()
    
    # 設定中文字型
    doc.styles['Normal'].font.name = '微軟正黑體'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微軟正黑體')
    
    # === 封面 ===
    title = doc.add_heading('問卷描述性統計報告', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitle = doc.add_paragraph('未上市櫃公司治理問卷分析')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = Pt(16)
    
    doc.add_paragraph()
    date_para = doc.add_paragraph(f'報告日期：{datetime.now().strftime("%Y年%m月%d日")}')
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_page_break()
    
    # === 1. 問卷分析 ===
    add_heading_with_style(doc, '一、問卷分析', level=1)
    
    p1 = doc.add_paragraph(
        '本計畫旨在深入探討未上市（櫃）公司在公司治理方面的現況，調查範圍涵蓋未上市（櫃）'
        '公司及其投資人，涉及多個產業領域與三個不同發展階段。透過多元且豐富的基礎資料，'
        '期望為不同階段的未上市（櫃）公司在公司治理方面提供具體洞察與實務建議。'
    )
    
    # === 2. 樣本介紹 ===
    add_heading_with_style(doc, '二、樣本介紹', level=1)
    
    doc.add_paragraph('本計畫樣本主要來自於臺灣未上市（櫃）公司及投資機構，樣本選取標準如下：')
    
    # 未上市櫃公司
    add_heading_with_style(doc, '(一) 未上市（櫃）公司', level=2)
    doc.add_paragraph(
        '本計畫根據國發基金投資方案的階段別差異，樣本公司依投資階段分為三類：'
        '第一階段選擇創業天使投資方案；第二階段選擇加強投資中小企業實施方案、'
        '加強投資策略性服務業實施方案、加強投資策略性製造業實施方案、'
        '加強投資文化創意產業實施方案；第三階段則選擇直接投資方案之未上市（櫃）公司。'
        '除此之外，本計畫亦挑選具有創新性及成長潛力的未上市（櫃）公司，'
        '涵蓋半導體、人工智慧、通訊、電子商務、生技、與機械製造等多元產業。'
    )
    
    # 投資人
    add_heading_with_style(doc, '(二) 投資人', level=2)
    doc.add_paragraph(
        '本計畫根據投資人所投資的未上市（櫃）公司發展階段，同樣設計三階段的問卷進行填答。'
        '參與的投資人背景涵蓋國內外知名創投公司與企業創投（Corporate Venture Capital, CVC）。'
        '其中，多數創投機構專注於第一階段與第二階段的未上市（櫃）公司投資，'
        '而企業創投則主要聚焦於第三階段的投資標的。'
    )
    
    # 樣本限制
    add_heading_with_style(doc, '(三) 樣本限制', level=2)
    doc.add_paragraph(
        '由於資料來源主要集中於臺灣，因此可能無法全面代表其他地區的未上市（櫃）公司與'
        '投資人特徵，存在一定的樣本偏誤。'
    )
    
    # === 3. 統計分析 ===
    add_heading_with_style(doc, '三、統計分析', level=1)
    
    # 基本統計
    total_samples = len(df)
    if 'respondent_type' in df.columns:
        company_count = len(df[df['respondent_type'] == '公司方'])
        investor_count = len(df[df['respondent_type'] == '投資方'])
    else:
        company_count = investor_count = 0
    
    doc.add_paragraph(f'本研究問卷共蒐集樣本 {total_samples} 份，樣本共分為兩類：')
    doc.add_paragraph(f'• 公司方：共計 {company_count} 份，涵蓋創辦人、高階主管及公司治理相關人員等')
    doc.add_paragraph(f'• 投資方：共計 {investor_count} 份，涵蓋創投、天使投資人及其他機構投資人等')
    
    # 樣本分佈表
    if 'respondent_type' in df.columns:
        sample_data = {
            'columns': ['類別', '份數', '百分比'],
            'data': [
                ['公司方', company_count, f'{company_count/total_samples*100:.1f}%'],
                ['投資方', investor_count, f'{investor_count/total_samples*100:.1f}%'],
                ['合計', total_samples, '100.0%']
            ]
        }
        add_statistics_table(doc, sample_data, title="表 1: 樣本分佈統計")
    
    # === 核心議題說明 ===
    add_heading_with_style(doc, '(一) 核心議題選定說明', level=2)
    
    doc.add_paragraph(
        '本報告從問卷中精選 20 個核心議題進行深入分析，涵蓋未上市櫃公司治理的六大關鍵面向。'
        '這些議題的選擇係基於公司治理理論與實務的重要性，以及對投資人決策與公司永續發展的影響程度。'
    )
    
    doc.add_paragraph()
    doc.add_paragraph('【六大關鍵面向】', style='Heading 4')
    
    dimension_para = doc.add_paragraph()
    dimension_para.add_run('1. 股權結構與控制權（2題）：').bold = True
    dimension_para.add_run('股權集中度與經營團隊持股是公司治理的基礎，直接影響決策效率與代理問題。')
    
    dimension_para = doc.add_paragraph()
    dimension_para.add_run('2. 股東會治理（3題）：').bold = True
    dimension_para.add_run('股東會的通知時效、決議記錄與董事出席情況，是保障股東權益的基本機制。')
    
    dimension_para = doc.add_paragraph()
    dimension_para.add_run('3. 董事會治理機制（5題）：').bold = True
    dimension_para.add_run('董事會的議程通知、決議記錄、召集程序、議事內容與開會頻率，是確保董事會有效運作的核心要素。')
    
    dimension_para = doc.add_paragraph()
    dimension_para.add_run('4. 財務報告與資訊透明度（5題）：').bold = True
    dimension_para.add_run('外部專業諮詢、財務查核、股權揭露、業務報告與財務報告頻率，是投資人監督管理層的重要基礎。')
    
    dimension_para = doc.add_paragraph()
    dimension_para.add_run('5. 內部控制與風險管理（3題）：').bold = True
    dimension_para.add_run('財務職能分工、財務紀錄處理與智慧財產權保護，是未上市櫃公司內部控制的基本要求。')
    
    dimension_para = doc.add_paragraph()
    dimension_para.add_run('6. 利害關係人治理（2題）：').bold = True
    dimension_para.add_run('員工激勵制度與利害關係人溝通機制，反映公司對人力資本與多方協作的重視程度。')
    
    doc.add_paragraph()
    doc.add_paragraph(
        '以下分別針對各項核心議題，提供完整的統計分析（包含公司方與投資方比較、'
        '公司發展階段分析）、統計檢定結果，以及業務意涵解讀。即使統計檢定未達顯著水準，'
        '仍透過描述性統計提供實務觀察，供政策制定與業務推動參考。'
    )
    
    doc.add_page_break()
    
    # === 議題分析 ===
    # 這裡需要根據實際欄位動態生成
    # 先返回基礎文件結構
    
    return doc

def add_topic_analysis(doc, df, topic_col, topic_title, topic_description):
    """
    新增單一議題的完整分析
    包含：標題、描述、表格、圖表、統計檢定、業務解讀
    即使統計檢定沒過也提供詳細敘述
    
    注意：如果df沒有'respondent_type'欄位，則只做整體分析，不做公司方vs投資方比較
    """
    
    add_heading_with_style(doc, topic_title, level=2)
    
    if topic_description:
        para = doc.add_paragraph(topic_description)
        para.runs[0].font.size = Pt(11)
    
    # (一) 公司方與投資方比較（如果有respondent_type欄位）
    if 'respondent_type' in df.columns:
        add_heading_with_style(doc, '(一) 公司方與投資方比較', level=3)
    else:
        add_heading_with_style(doc, '(一) 整體分佈情況', level=3)
    
    if topic_col not in df.columns:
        doc.add_paragraph('本題目不存在於資料中。')
        return doc
    
    # 清理資料
    if 'respondent_type' in df.columns:
        df_clean = df[[topic_col, 'respondent_type']].dropna()
    else:
        df_clean = df[[topic_col]].dropna()
    
    if len(df_clean) == 0:
        doc.add_paragraph('本題目無有效樣本資料。')
        return doc
    
    # 計算分佈
    if 'respondent_type' in df.columns:
        crosstab = pd.crosstab(df_clean[topic_col], df_clean['respondent_type'], margins=True)
        crosstab_pct = pd.crosstab(df_clean[topic_col], df_clean['respondent_type'], normalize='columns') * 100
    else:
        # 只計算整體分佈
        value_counts = df_clean[topic_col].value_counts()
        total_count = len(df_clean)
        crosstab = pd.DataFrame({
            '次數': value_counts,
            '百分比': (value_counts / total_count * 100).round(1)
        })
        crosstab.loc['合計'] = [total_count, 100.0]
    
    # 生成表格
    if 'respondent_type' in df.columns:
        table_data = {
            'columns': ['選項', '公司方人數', '公司方百分比', '投資方人數', '投資方百分比', '合計'],
            'data': []
        }
        
        for idx in crosstab.index:
            if idx != 'All':
                company_count = crosstab.loc[idx, '公司方'] if '公司方' in crosstab.columns else 0
                company_pct = f"{crosstab_pct.loc[idx, '公司方']:.1f}%" if '公司方' in crosstab.columns else '-'
                investor_count = crosstab.loc[idx, '投資方'] if '投資方' in crosstab.columns else 0
                investor_pct = f"{crosstab_pct.loc[idx, '投資方']:.1f}%" if '投資方' in crosstab.columns else '-'
                total = crosstab.loc[idx, 'All']
                
                table_data['data'].append([
                    str(idx), 
                    company_count, 
                    company_pct, 
                    investor_count, 
                    investor_pct, 
                    total
                ])
        
        # 合計行
        company_total = crosstab.loc['All', '公司方'] if '公司方' in crosstab.columns else 0
        investor_total = crosstab.loc['All', '投資方'] if '投資方' in crosstab.columns else 0
        table_data['data'].append([
            '合計',
            company_total,
            '100.0%',
            investor_total,
            '100.0%',
            crosstab.loc['All', 'All']
        ])
        
        add_statistics_table(doc, table_data, title=f"{topic_title} - 受訪者類型分佈表")
        
        # === 加入長條圖 ===
        doc.add_paragraph()
        doc.add_paragraph('【圖表呈現】', style='Heading 4')
        
        try:
            # 獲取所有類別（排除 'All'）
            categories = [idx for idx in crosstab.index if idx != 'All']
            
            # 創建長條圖
            chart_title = f"{topic_title} - 公司方與投資方比較"
            fig = create_bar_chart(crosstab, crosstab_pct, chart_title, categories)
            
            # 儲存圖片
            chart_filename = f"/tmp/chart_{hash(topic_title)}.png"
            if save_plotly_as_image(fig, chart_filename):
                # 將圖片插入 Word 文件
                doc.add_picture(chart_filename, width=Inches(6))
                # 圖片置中
                last_paragraph = doc.paragraphs[-1]
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph()
                
                # 清理臨時檔案
                try:
                    os.remove(chart_filename)
                except:
                    pass
            else:
                doc.add_paragraph('（圖表生成失敗）')
        except Exception as e:
            print(f"圖表插入失敗: {e}")
            doc.add_paragraph(f'（圖表生成時發生錯誤）')
        
        # 統計檢定
        chi_result = calculate_chi_square(df_clean, topic_col, 'respondent_type')
    else:
        # 只有整體分佈，沒有公司方vs投資方比較
        table_data = {
            'columns': ['選項', '次數', '百分比'],
            'data': []
        }
        
        for idx in crosstab.index:
            if idx != '合計':
                table_data['data'].append([
                    str(idx),
                    int(crosstab.loc[idx, '次數']),
                    f"{crosstab.loc[idx, '百分比']:.1f}%"
                ])
        
        # 合計行
        table_data['data'].append([
            '合計',
            int(crosstab.loc['合計', '次數']),
            f"{crosstab.loc['合計', '百分比']:.1f}%"
        ])
        
        add_statistics_table(doc, table_data, title=f"{topic_title} - 整體分佈表")
        
        # === 加入長條圖 ===
        doc.add_paragraph()
        doc.add_paragraph('【圖表呈現】', style='Heading 4')
        
        try:
            # 獲取所有類別（排除 '合計'）
            categories = [idx for idx in crosstab.index if idx != '合計']
            percentages = [crosstab.loc[idx, '百分比'] for idx in categories]
            
            # 智慧排序
            sorted_categories = smart_sort_categories(categories)
            sorted_percentages = [crosstab.loc[cat, '百分比'] for cat in sorted_categories]
            
            # 創建長條圖（使用與 cloud_app.py 一致的樣式）
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=sorted_categories,
                y=sorted_percentages,
                marker_color='#1f77b4',
                text=[f"{p:.1f}%" for p in sorted_percentages],
                textposition='auto'
            ))
            
            fig.update_layout(
                title='整體選項分佈',
                xaxis_title='選項',
                yaxis_title='比例 (%)',
                template='plotly_white',
                height=500,
                xaxis_tickangle=-45,
                xaxis={'categoryorder': 'array', 'categoryarray': sorted_categories},
                showlegend=False,
                font=dict(family='Microsoft JhengHei', size=12)
            )
            
            # 儲存圖片
            chart_filename = f"/tmp/chart_{hash(topic_title)}.png"
            if save_plotly_as_image(fig, chart_filename):
                # 將圖片插入 Word 文件
                doc.add_picture(chart_filename, width=Inches(5))
                # 圖片置中
                last_paragraph = doc.paragraphs[-1]
                last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph()
                
                # 清理臨時檔案
                try:
                    os.remove(chart_filename)
                except:
                    pass
            else:
                doc.add_paragraph('（圖表生成失敗）')
        except Exception as e:
            print(f"圖表插入失敗: {e}")
            doc.add_paragraph(f'（圖表生成時發生錯誤）')
        
        # 無法做公司方vs投資方的統計檢定
        chi_result = None
    
    # 統計檢定結果顯示
    if 'respondent_type' in df.columns:
        
        doc.add_paragraph('【統計檢定】', style='Heading 4')
        
        if chi_result and chi_result['p_value'] is not None:
            p = chi_result['p_value']
            
            significance_level = ''
            if p < 0.001:
                significance_level = '***（p < 0.001，極顯著）'
            elif p < 0.01:
                significance_level = '**（p < 0.01，高度顯著）'
            elif p < 0.05:
                significance_level = '*（p < 0.05，顯著）'
            else:
                significance_level = 'n.s.（p ≥ 0.05，無顯著差異）'
            
            doc.add_paragraph(f"檢定方法：{chi_result['method']}")
            if 'statistic' in chi_result:
                doc.add_paragraph(f"檢定統計量：{chi_result['statistic']:.3f}")
            if 'dof' in chi_result:
                doc.add_paragraph(f"自由度：{chi_result['dof']}")
            doc.add_paragraph(f"顯著性水準：p = {p:.4f} {significance_level}")
        else:
            doc.add_paragraph("由於樣本數不足或資料特性，無法進行統計檢定。")
        
        # 業務解讀（即使檢定沒過也提供）
        doc.add_paragraph()
        doc.add_paragraph('【業務意涵與解讀】', style='Heading 4')
        
        # 計算各類別最高佔比
        if '公司方' in crosstab.columns and '投資方' in crosstab.columns:
            company_top = crosstab_pct['公司方'].idxmax()
            investor_top = crosstab_pct['投資方'].idxmax()
            company_top_pct = crosstab_pct.loc[company_top, '公司方']
            investor_top_pct = crosstab_pct.loc[investor_top, '投資方']
            
            if chi_result and chi_result['p_value'] < 0.05:
                # 有顯著差異
                doc.add_paragraph(
                    f"統計檢定顯示公司方與投資方在本議題上存在顯著差異（p = {chi_result['p_value']:.4f}）。"
                    f"從數據分佈來看，公司方最多選擇「{company_top}」（{company_top_pct:.1f}%），"
                    f"而投資方則以「{investor_top}」為主（{investor_top_pct:.1f}%）。"
                    f"此一差異反映出雙方在此議題上的認知或實務存在明顯落差，建議進一步深入探討造成差異的原因，"
                    f"以利於公司與投資方之間建立更有效的溝通機制。"
                )
            else:
                # 無顯著差異或無法檢定
                if chi_result:
                    doc.add_paragraph(
                        f"統計檢定顯示公司方與投資方在本議題上無顯著差異（p = {chi_result['p_value']:.4f}）。"
                    )
                else:
                    doc.add_paragraph("本議題樣本數較少，統計檢定結果僅供參考。")
                
                doc.add_paragraph(
                    f"從描述性統計來看，公司方最多選擇「{company_top}」（{company_top_pct:.1f}%），"
                    f"投資方最多選擇「{investor_top}」（{investor_top_pct:.1f}%）。"
                    f"雖然統計上未達顯著差異，但此一數據分佈顯示雙方在此議題上具有相似的認知或做法。"
                    f"此結果有助於了解當前未上市櫃公司治理的共識點，可作為推動相關政策或實務的基礎。"
                )
    
    # (二) 公司階段分析
    if 'phase' in df.columns and topic_col in df.columns:
        add_heading_with_style(doc, '(二) 公司發展階段分析', level=3)
        
        # 清理資料
        df_phase = df[[topic_col, 'phase']].dropna()
        
        if len(df_phase) > 0:
            # 階段交叉表
            phase_crosstab = pd.crosstab(df_phase[topic_col], df_phase['phase'], margins=True)
            phase_crosstab_pct = pd.crosstab(df_phase[topic_col], df_phase['phase'], normalize='columns') * 100
            
            # 生成階段表格
            phases = sorted([p for p in phase_crosstab.columns if p != 'All'])
            
            if len(phases) > 0:
                table_columns = ['選項']
                for phase in phases:
                    table_columns.extend([f'{phase}人數', f'{phase}百分比'])
                table_columns.append('合計')
                
                table_data = {
                    'columns': table_columns,
                    'data': []
                }
                
                for idx in phase_crosstab.index:
                    if idx != 'All':
                        row = [str(idx)]
                        for phase in phases:
                            if phase in phase_crosstab.columns:
                                count = phase_crosstab.loc[idx, phase]
                                pct = f"{phase_crosstab_pct.loc[idx, phase]:.1f}%"
                            else:
                                count = 0
                                pct = '-'
                            row.extend([count, pct])
                        row.append(phase_crosstab.loc[idx, 'All'])
                        table_data['data'].append(row)
                
                # 合計行
                total_row = ['合計']
                for phase in phases:
                    if phase in phase_crosstab.columns:
                        total_row.extend([phase_crosstab.loc['All', phase], '100.0%'])
                    else:
                        total_row.extend([0, '-'])
                total_row.append(phase_crosstab.loc['All', 'All'])
                table_data['data'].append(total_row)
                
                add_statistics_table(doc, table_data, title=f"{topic_title} - 公司發展階段分佈表")
                
                # === 加入階段比較長條圖 ===
                doc.add_paragraph()
                doc.add_paragraph('【圖表呈現】', style='Heading 4')
                
                try:
                    # 獲取所有類別（排除 'All'）
                    categories = [idx for idx in phase_crosstab.index if idx != 'All']
                    
                    # 創建階段比較長條圖
                    chart_title = f"{topic_title} - 公司發展階段比較"
                    fig = create_phase_chart(phase_crosstab, phase_crosstab_pct, chart_title, categories, phases)
                    
                    # 儲存圖片
                    chart_filename = f"/tmp/phase_chart_{hash(topic_title)}.png"
                    if save_plotly_as_image(fig, chart_filename):
                        # 將圖片插入 Word 文件
                        doc.add_picture(chart_filename, width=Inches(6))
                        # 圖片置中
                        last_paragraph = doc.paragraphs[-1]
                        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        doc.add_paragraph()
                        
                        # 清理臨時檔案
                        try:
                            os.remove(chart_filename)
                        except:
                            pass
                    else:
                        doc.add_paragraph('（圖表生成失敗）')
                except Exception as e:
                    print(f"階段圖表插入失敗: {e}")
                    doc.add_paragraph(f'（圖表生成時發生錯誤）')
                
                # Kruskal-Wallis 檢定
                doc.add_paragraph('【統計檢定】', style='Heading 4')
                
                try:
                    # 嘗試將選項轉為數值進行檢定
                    phase_groups = [df_phase[df_phase['phase'] == p][topic_col].dropna() for p in phases]
                    valid_groups = [g for g in phase_groups if len(g) > 0]
                    
                    if len(valid_groups) >= 2 and all(len(g) >= 3 for g in valid_groups):
                        # 嘗試轉換為數值
                        numeric_groups = []
                        for g in valid_groups:
                            try:
                                numeric_groups.append(pd.to_numeric(g, errors='coerce').dropna())
                            except:
                                pass
                        
                        if len(numeric_groups) >= 2 and all(len(g) >= 3 for g in numeric_groups):
                            H_stat, p_val = kruskal(*numeric_groups)
                            
                            significance = ''
                            if p_val < 0.05:
                                significance = '*（顯著）'
                            else:
                                significance = 'n.s.（無顯著差異）'
                            
                            doc.add_paragraph(f"檢定方法：Kruskal-Wallis H 檢定（無母數檢定）")
                            doc.add_paragraph(f"H 統計量：{H_stat:.3f}")
                            doc.add_paragraph(f"顯著性水準：p = {p_val:.4f} {significance}")
                            
                            # 階段分析解讀
                            doc.add_paragraph()
                            doc.add_paragraph('【階段差異分析】', style='Heading 4')
                            
                            if p_val < 0.05:
                                doc.add_paragraph(
                                    f"統計檢定顯示不同發展階段的公司在本議題上存在顯著差異（p = {p_val:.4f}）。"
                                    f"此結果表明公司發展階段確實影響此議題的表現或認知。"
                                    f"建議針對不同階段公司的特性，提供差異化的治理建議或輔導措施。"
                                )
                            else:
                                doc.add_paragraph(
                                    f"統計檢定顯示不同發展階段的公司在本議題上無顯著差異（p = {p_val:.4f}）。"
                                    f"此結果表明本議題可能是跨階段的共同關注點，不因公司發展階段而有明顯變化。"
                                )
                        else:
                            raise ValueError("資料無法轉換為數值")
                    else:
                        raise ValueError("各組樣本數不足")
                        
                except Exception as e:
                    doc.add_paragraph("由於資料類型為類別變項或樣本數不足，無法進行階段差異檢定。")
                    doc.add_paragraph()
                    doc.add_paragraph('【階段分佈觀察】', style='Heading 4')
                    
                    # 提供描述性觀察
                    phase_descriptions = []
                    for phase in phases:
                        if phase in phase_crosstab_pct.columns:
                            top_option = phase_crosstab_pct[phase].idxmax()
                            top_pct = phase_crosstab_pct.loc[top_option, phase]
                            phase_descriptions.append(f"{phase}主要選擇「{top_option}」（{top_pct:.1f}%）")
                    
                    if phase_descriptions:
                        doc.add_paragraph(
                            f"從各階段的分佈觀察：{'；'.join(phase_descriptions)}。"
                            f"雖無法進行統計檢定，但透過描述性統計仍可觀察不同階段公司的選擇傾向，"
                            f"提供未來政策制定或業務推動的參考。"
                        )
        else:
            doc.add_paragraph('本題目無有效的階段資料。')
    
    doc.add_paragraph()  # 空行
    doc.add_page_break()  # 每個議題後分頁
    return doc

def generate_full_descriptive_report(df, output_path="/workspaces/work1/問卷描述性統計報告_完整版.docx"):
    """
    生成完整描述性統計報告（Word 格式）
    包含更多題目，附上政府統計風格表格
    即使統計檢定沒過也提供敘述性分析
    """
    print("開始生成描述性統計報告...")
    
    # 創建基礎文件
    doc = generate_descriptive_report_word(df, output_path)
    
    # 定義要分析的議題（擴展到20個核心重要議題）
    # 涵蓋：股權結構、董事會治理、資訊揭露、內部控制、利害關係人等五大面向
    # 欄位名稱已根據實際CSV檔案調整
    topics = [
        # === 一、股權結構與控制權（2題）===
        {
            'col': '請問公司大股東（持股5%以上）合計持股比例為多少？',
            'title': '3.1 大股東合計持股比例',
            'description': '分析大股東合計持股比例，評估股權集中程度與控制權分配。股權集中度是公司治理的基礎指標，影響決策效率與股東權益保護。'
        },
        {
            'col': '請問公司經營團隊合計持股比例為多少？',
            'title': '3.2 經營團隊持股比例',
            'description': '分析經營團隊持股情況，了解管理層與公司利益的一致性程度。經營團隊持股比例反映所有權與經營權的結合程度，是評估代理問題的關鍵指標。'
        },
        
        # === 二、股東會治理（3題）===
        {
            'col': '股東會結構與運作 - 公司股東常會的議程及相關資料能在20天前通知，並以可存證的方式（如：掛號或經股東同意的電子方式）寄發',
            'title': '3.3 股東會通知時效性',
            'description': '評估股東會召開前的資訊揭露時效。充分的準備時間讓股東能審慎評估議案，是保障股東權益的基本要求。'
        },
        {
            'col': '股東會結構與運作 - 公司股東會決議方式能夠清楚載明，且議事錄完整記載會議資訊（如時間、地點、主席、決議方法及結果）',
            'title': '3.4 股東會決議與議事錄完整性',
            'description': '評估股東會決策的透明度與記錄完整性。完整的議事錄是確保決策可追溯性的基礎，也是公司治理品質的重要指標。'
        },
        {
            'col': '股東會結構與運作 - 公司董事長及董事通常能夠親自出席股東常會',
            'title': '3.5 董事出席股東會情況',
            'description': '調查董事對股東會的重視程度。董事親自出席能直接回應股東關切，展現負責任的治理態度。'
        },
        
        # === 三、董事會治理機制（5題）===
        {
            'col': '董事會結構與運作 - 公司董事會的議程及相關資料能在3天前通知，並以掛號或電子方式（經過股東同意或於公司章程中載明）寄發',
            'title': '3.6 董事會通知時效性',
            'description': '評估董事會會前資訊揭露的及時性。提前通知讓董事有充分時間準備，提升會議品質與決策效率。'
        },
        {
            'col': '董事會結構與運作 - 公司董事會決議方式能夠清楚載明，且議事錄完整記載會議資訊（如時間、地點、主席、決議方法及結果）',
            'title': '3.7 董事會決議與議事錄完整性',
            'description': '評估董事會決策的透明度與記錄完整性。完整的議事錄是確保董事會決策可追溯性的基礎，也是公司治理品質的重要指標。'
        },
        {
            'col': '董事會結構與運作 - 公司通常每年召開一次股東常會，並是由董事會召集',
            'title': '3.8 股東會召集程序',
            'description': '調查股東會的召集程序與規律性。定期召開股東會並由董事會召集，是公司治理的基本程序要求。'
        },
        {
            'col': '公司定期性董事會的議事內容，通常包含以下哪些項目？ (可複選)',
            'title': '3.9 董事會議事內容廣度',
            'description': '評估董事會討論議題的完整性。多元的議事內容反映董事會對公司營運的全方位監督。'
        },
        {
            'col': '董事會結構與運作 - 在過去12個月內，貴公司董事會的召開頻率為何？',
            'title': '3.10 董事會召開頻率',
            'description': '調查董事會的開會頻率。定期召開董事會是確保董事會有效運作、即時監督公司營運的基本要求。'
        },
        
        # === 四、財務報告與資訊透明度（5題）===
        {
            'col': '董事會結構與運作 - 公司諮詢顧問／業師的頻率為何？',
            'title': '3.11 外部專業諮詢頻率',
            'description': '評估公司尋求外部專業意見的積極程度。適度的外部諮詢能引入專業觀點，提升決策品質。'
        },
        {
            'col': '公司最近一期的年度財務報表，是否委任外部會計師進行查核簽證？若有，其遵循的會計準則為何？',
            'title': '3.12 財務報表查核與會計準則',
            'description': '調查公司財務報表的外部查核情況與會計準則遵循。外部查核是確保財務資訊可信度的關鍵機制，是投資人決策的重要依據。'
        },
        {
            'col': '資訊透明度 - 公司通常多久向股東揭露董事、監察人、經理人及持股超過10%大股東的持股情形、股權質押比率與前十大股東之股權結構圖或表',
            'title': '3.13 股權結構資訊揭露',
            'description': '評估公司股權結構資訊的透明度與更新頻率。股權結構揭露讓投資人了解公司控制權變動與潛在利益衝突，是重大資訊透明的體現。'
        },
        {
            'col': '資訊透明度 - 公司通常多久向股東提供業務報告（如營運、研發進度等）',
            'title': '3.14 業務報告提供頻率',
            'description': '了解公司向股東揭露業務資訊的頻率。定期的業務報告讓股東掌握公司營運狀況，是資訊透明的重要體現。'
        },
        {
            'col': '資訊透明度 - 公司通常多久向股東提供財務報告',
            'title': '3.15 財務報告提供頻率',
            'description': '了解公司向股東揭露財務資訊的頻率與即時性。定期且及時的財務報告是股東監督管理層的基礎，反映公司資訊透明度。'
        },
        
        # === 五、內部控制與風險管理（3題）===
        {
            'col': '內控與風險評估（含財務與營運風險） - 公司由不同人員分別負責出納與會計',
            'title': '3.16 財務職能分工',
            'description': '評估公司財務內部控制的基本分工情況。出納與會計分工是財務內控的基石，有效防止舞弊與錯誤，是未上市櫃公司最基本但最關鍵的控制點。'
        },
        {
            'col': '內控與風險評估（含財務與營運風險） - 公司財務紀錄由專責人員或外部會計師協助處理',
            'title': '3.17 財務紀錄專業處理',
            'description': '調查公司財務紀錄的專業處理機制。專責人員或外部專業協助能確保財務紀錄的準確性與合規性。'
        },
        {
            'col': '內控與風險評估（含財務與營運風險） - 公司開發的專利、商標等智慧財產權，均已登記在公司名下',
            'title': '3.18 智慧財產權保護',
            'description': '評估公司對智慧財產權的保護措施。完整的智財權登記能保障公司核心資產，防範法律風險。'
        },
        
        # === 六、利害關係人治理（2題）===
        {
            'col': '利害關係人 - 公司員工分紅制度設計能有效激勵員工',
            'title': '3.19 員工激勵制度',
            'description': '評估公司員工激勵機制的有效性。有效的員工激勵制度能將員工利益與公司長期發展結合，是人力資本管理與公司永續經營的關鍵。'
        },
        {
            'col': '利害關係人 - 公司已建立與主要利害關係人（如員工、債權人、外部投資人等）的溝通管道',
            'title': '3.20 利害關係人溝通機制',
            'description': '調查公司與利害關係人的溝通管道建立情況。良好的溝通機制能促進多方協作，降低衝突風險，是企業永續經營的基礎。'
        },
    ]
    
    # 逐題分析
    analyzed_count = 0
    for topic in topics:
        if topic['col'] in df.columns:
            print(f"正在分析: {topic['title']}")
            try:
                add_topic_analysis(doc, df, topic['col'], topic['title'], topic['description'])
                analyzed_count += 1
            except Exception as e:
                print(f"分析 {topic['title']} 時發生錯誤: {e}")
                doc.add_paragraph(f"[{topic['title']} 資料不足或分析發生錯誤]")
        else:
            print(f"欄位不存在，跳過: {topic['col']}")
    
    print(f"✅ 共分析 {analyzed_count} 個議題")
    
    # 儲存文件
    doc.save(output_path)
    print(f"✅ 報告已儲存至: {output_path}")
    
    return output_path

if __name__ == "__main__":
    # 測試用
    print("描述性統計報告生成器已就緒")
