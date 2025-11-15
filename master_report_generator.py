"""
Master report generator
- 讀取並合併 7 個 STANDARD_... CSV
- 新增 respondent_type 與 phase 欄位
- 數據標準化（簡單處理函式）
- 呼叫 descriptive_report_generator.generate_descriptive_report_word 產生 Word

用法：
    python3 master_report_generator.py

注意：預設會在 DRY_RUN 模式下執行（不會寫圖檔或 docx），要產出真實檔案請先設定環境變數 DRY_RUN="0" 或移除。"""

import os
import pandas as pd
from descriptive_report_generator import generate_descriptive_report_word

# 要合併的檔名（workspace 內存在）
FILES = [
    'STANDARD_8RG8Y_未上市櫃公司治理問卷第一階段_202511050604_690ae8db08878.csv',
    'STANDARD_7RGxP_未上市櫃公司治理問卷第二階段_202511050605_690ae92a9a127.csv',
    'STANDARD_Yb9D2_未上市櫃公司治理問卷第三階段_202511050605_690ae9445a228.csv',
    'STANDARD_NwNYM_未上市櫃公司治理問卷第一階段投資方_202511060133_690bfaccec28e.csv',
    'STANDARD_v2xYO_未上市櫃公司治理問卷第二階段投資方_202511060133_690bfae9b9065.csv',
    'STANDARD_we89e_未上市櫃公司治理問卷第三階段投資方_202511060133_690bfb0524491.csv',
    'STANDARD_v2xkX_未上市櫃公司治理問卷_202511060532_690c3305c62b5.csv'
]

# 對應 respondent_type
COMPANY_FILES = [
    'STANDARD_8RG8Y', 'STANDARD_7RGxP', 'STANDARD_Yb9D2'
]
INVESTOR_FILES = [
    'STANDARD_NwNYM', 'STANDARD_v2xYO', 'STANDARD_we89e'
]

# 根據檔名判斷 phase
# 以更穩健的方式從檔名推斷 phase（支援中文註記或代碼匹配）
PHASE_CODE_MAP = {
    '8RG8Y': '第一階段',
    '7RGxP': '第二階段',
    'Yb9D2': '第三階段',
    'NwNYM': '第一階段',
    'v2xYO': '第二階段',
    'we89e': '第三階段',
    'v2xkX': '第二階段'
}

def infer_role_from_filename(fname):
    base = os.path.basename(fname)
    for k in COMPANY_FILES:
        if k in base:
            return '公司方'
    for k in INVESTOR_FILES:
        if k in base:
            return '投資方'
    # fallback: look for Chinese keywords
    if '公司' in base and '投資' not in base:
        return '公司方'
    if '投資' in base or '投資方' in base:
        return '投資方'
    # default
    return '公司方'


def infer_phase_from_filename(fname):
    base = os.path.basename(fname)
    # 優先搜尋中文階段標記
    if '第一階段' in base or '第一階' in base or '第一' in base:
        return '第一階段'
    if '第二階段' in base or '第二階' in base or '第二' in base:
        return '第二階段'
    if '第三階段' in base or '第三階' in base or '第三' in base:
        return '第三階段'

    # 再以代碼匹配
    for code, phase in PHASE_CODE_MAP.items():
        if code in base:
            return phase

    # 其他情況回傳 None，讓後續流程以問卷欄位或預設處理
    return None


def normalize_answer(val):
    if pd.isna(val):
        return val
    s = str(val).strip()
    # 移除尾端的 人 字
    s = s.replace('人', '')
    # 移除不必要空白與特殊符號
    s = s.replace('\u3000', ' ').strip()
    # 統一百分比範例格式
    s = s.replace('％', '%')
    s = s.replace(' ', '')
    return s


def load_and_tag(filepath):
    df = pd.read_csv(filepath, dtype=str)
    base = os.path.basename(filepath)
    # 保留來源檔名以便後續判斷與調查
    df['_source_file'] = base

    # 判斷 respondent_type 與 phase（從檔名推斷）
    tag = infer_role_from_filename(base)
    df['respondent_type'] = tag

    phase = infer_phase_from_filename(base)
    df['phase'] = phase

    # normalize answers
    for col in df.columns:
        if col in ['respondent_type', 'phase']:
            continue
        df[col] = df[col].apply(normalize_answer)

    return df


def main():
    dfs = []
    for f in FILES:
        if not os.path.exists(f):
            print(f"警告：找不到檔案 {f}，已跳過")
            continue
        try:
            df = load_and_tag(f)
            dfs.append(df)
            print(f"已載入: {f} ({len(df)} 列)")
        except Exception as e:
            print(f"載入 {f} 時發生錯誤: {e}")

    if not dfs:
        raise SystemExit('沒有可合併的檔案，請確認 CSV 檔是否存在')

    master = pd.concat(dfs, ignore_index=True, sort=False)
    print(f"合併後總筆數: {len(master)}")

    # 呼叫報表生成器
    output = generate_descriptive_report_word(master, output_filename='問卷描述性統計報告_Master.docx')
    print('生成完成: ', output)


if __name__ == '__main__':
    main()
