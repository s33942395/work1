"""
PDF 分析工具 - 讀取桶.pdf並分析統計報告格式
"""
import pdfplumber
import re

def extract_pdf_text(pdf_path):
    """提取 PDF 文字內容"""
    text_content = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            
            print(f"PDF 總頁數: {total_pages}")
            print("="*80)
            
            for page_num in range(total_pages):
                page = pdf.pages[page_num]
                text = page.extract_text()
                text_content.append({
                    'page': page_num + 1,
                    'text': text
                })
                
                print(f"\n[第 {page_num + 1} 頁]")
                print("-"*80)
                print(text[:500] if len(text) > 500 else text)  # 顯示前500字元
                
    except Exception as e:
        print(f"讀取 PDF 時發生錯誤: {e}")
        return None
    
    return text_content

def analyze_report_structure(text_content):
    """分析報告結構"""
    if not text_content:
        return
    
    print("\n\n" + "="*80)
    print("報告結構分析")
    print("="*80)
    
    # 合併所有頁面文字
    full_text = "\n".join([page['text'] for page in text_content])
    
    # 尋找章節標題 (可能的模式)
    section_patterns = [
        r'第[一二三四五六七八九十\d]+章.*',
        r'[一二三四五六七八九十]+、.*',
        r'\d+\.\s*[^\n]+',
        r'壹|貳|參|肆|伍|陸|柒|捌|玖|拾',
    ]
    
    print("\n檢測到的章節結構:")
    for pattern in section_patterns:
        matches = re.findall(pattern, full_text, re.MULTILINE)
        if matches:
            print(f"\n使用模式: {pattern}")
            for match in matches[:10]:  # 只顯示前10個
                print(f"  - {match}")
    
    # 尋找統計相關關鍵字
    print("\n\n統計方法關鍵字:")
    stats_keywords = [
        '平均數', '標準差', '變異數', '中位數', '眾數',
        '卡方檢定', 'Chi-square', 'χ²', 
        't檢定', 't-test',
        'ANOVA', '變異數分析',
        '相關係數', '迴歸分析',
        'p值', 'p-value', '顯著性',
        '信度', 'Cronbach', 'α係數',
        '效度', '因素分析',
        '樣本數', '問卷', '量表'
    ]
    
    found_keywords = {}
    for keyword in stats_keywords:
        count = full_text.count(keyword)
        if count > 0:
            found_keywords[keyword] = count
    
    for keyword, count in sorted(found_keywords.items(), key=lambda x: x[1], reverse=True):
        print(f"  {keyword}: 出現 {count} 次")
    
    # 尋找圖表
    print("\n\n圖表相關:")
    chart_keywords = ['圖', '表', 'Figure', 'Table', '圖表']
    for keyword in chart_keywords:
        matches = re.findall(f'{keyword}[\\s]*\\d+[^\\n]*', full_text)
        if matches:
            print(f"\n{keyword}的引用:")
            for match in matches[:5]:  # 顯示前5個
                print(f"  - {match}")

def save_full_text(text_content, output_file='pdf_content.txt'):
    """將完整內容儲存到文字檔"""
    if not text_content:
        return
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            for page in text_content:
                f.write(f"\n{'='*80}\n")
                f.write(f"第 {page['page']} 頁\n")
                f.write(f"{'='*80}\n")
                f.write(page['text'])
                f.write("\n\n")
        
        print(f"\n\n完整內容已儲存至: {output_file}")
    except Exception as e:
        print(f"儲存文字檔時發生錯誤: {e}")

if __name__ == "__main__":
    pdf_path = "/workspaces/work1/桶.pdf"
    
    print("開始分析 PDF...")
    text_content = extract_pdf_text(pdf_path)
    
    if text_content:
        analyze_report_structure(text_content)
        save_full_text(text_content)
        print("\n分析完成!")
    else:
        print("無法讀取 PDF 內容")
