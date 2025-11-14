#!/usr/bin/env python3
"""
生成測試報告以驗證所有功能
"""
import sys
from descriptive_report_generator import generate_full_descriptive_report

# CSV檔案路徑
csv_files = {
    '第一階段': '/workspaces/work1/STANDARD_8RG8Y_未上市櫃公司治理問卷第一階段_202511050604_690ae8db08878.csv',
    '第一階段投資方': '/workspaces/work1/STANDARD_NwNYM_未上市櫃公司治理問卷第一階段投資方_202511060133_690bfaccec28e.csv',
    '第二階段': '/workspaces/work1/STANDARD_7RGxP_未上市櫃公司治理問卷第二階段_202511050605_690ae92a9a127.csv',
    '第二階段投資方': '/workspaces/work1/STANDARD_v2xYO_未上市櫃公司治理問卷第二階段投資方_202511060133_690bfae9b9065.csv',
    '第三階段': '/workspaces/work1/STANDARD_Yb9D2_未上市櫃公司治理問卷第三階段_202511050605_690ae9445a228.csv',
    '第三階段投資方': '/workspaces/work1/STANDARD_we89e_未上市櫃公司治理問卷第三階段投資方_202511060133_690bfb0524491.csv'
}

output_path = '/workspaces/work1/test_report_with_reliability.docx'

print("="*60)
print("開始生成完整測試報告")
print("="*60)
print(f"輸出檔案: {output_path}")
print()

import pandas as pd
all_dfs = []
for label, path in csv_files.items():
    import os
    if not os.path.exists(path):
        print(f"警告：CSV 檔案不存在，已跳過：{path}")
        continue
    try:
        df = pd.read_csv(path)
    except Exception as e:
        print(f"讀取 CSV 失敗（跳過）：{path} -> {e}")
        continue
    df['_source_file'] = label
    all_dfs.append(df)
df_merged = pd.concat(all_dfs, ignore_index=True)

try:
    result_path = generate_full_descriptive_report(
        df_merged,
        output_path=output_path
    )
    print()
    print("="*60)
    print("✅ 報告生成成功!")
    print("="*60)
    print(f"檔案位置: {result_path}")
    print()
    print("報告內容包含:")
    print("  ✓ 所有階段的描述性統計")
    print("  ✓ 公司方與投資方的數據對比")
    print("  ✓ 階段分布分析與數據解讀")
    print("  ✓ 信度與效度分析(Cronbach's Alpha, KMO, Bartlett)")
    sys.exit(0)
    
except Exception as e:
    print()
    print("="*60)
    print("❌ 報告生成失敗")
    print("="*60)
    print(f"錯誤訊息: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
