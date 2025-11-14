# 未上市櫃公司治理問卷分析系統

## 📊 專案簡介

這是一個基於 Streamlit 開發的問卷分析系統，專門用於分析未上市櫃公司治理問卷數據。系統提供智慧題目合併、統計分析、視覺化圖表等功能。

## ✨ 主要功能

### 1. 智慧題目去重與合併
- 自動偵測相似題目並合併
- 處理公司方與投資方的對應題目
- 支援複選題、數值題、類別題的智慧合併

### 2. 深度統計分析
- **公司方 vs 投資方比較**：卡方檢定、Mann-Whitney U 檢定、Fisher 精確檢定
- **階段比較分析**：比較一階段、二階段、三階段的差異
- 自動計算顯著性水準並標註

### 3. 視覺化圖表
- 堆疊長條圖、分組長條圖
- 盒狀圖（支援階段分析）
- 互動式圖表（Plotly）

### 4. 報告推薦系統
- 自動計算題目優先順序
- 推薦最適合寫入報告的題目
- 提供分析洞察與建議

## 🚀 部署方式

### Streamlit Cloud 部署

1. 將專案推送到 GitHub
2. 前往 [Streamlit Cloud](https://share.streamlit.io/)
3. 使用 GitHub 帳號登入
4. 點擊「New app」
5. 選擇此 repository 和 `cloud_app.py`
6. 點擊「Deploy」

### 本地運行

```bash
pip install -r requirements.txt
streamlit run cloud_app.py
```

## 📁 資料格式

系統支援以下 CSV 檔案格式：
- 公司方問卷（第一、二、三階段）
- 投資方問卷（第一、二、三階段）

CSV 檔案應包含：
- 題目欄位（問卷題目）
- respondent_type 欄位（公司方/投資方）
- phase 欄位（階段標記）

## 📦 技術堆疊

- **Frontend**: Streamlit
- **Data Processing**: Pandas, NumPy
- **Visualization**: Plotly
- **Statistics**: SciPy
- **Language**: Python 3.11+

## 📝 使用說明

1. **上傳檔案**：支援單一或多個 CSV 檔案上傳
2. **選擇分析模式**：
   - 合併分析：自動合併相似題目並進行統計分析
   - 逐題瀏覽：查看所有原始題目
3. **查看推薦題目**：系統自動推薦具有統計意義的題目
4. **深度分析**：選擇題目進行公司方/投資方、階段比較分析

## 🔧 配置說明

- `requirements.txt`: Python 套件依賴
- `.streamlit/config.toml`: Streamlit 配置
- `cloud_app.py`: 主程式

## 📄 授權

此專案僅供內部使用。

## 👥 聯絡資訊

如有問題請聯繫專案維護者。
<!-- Trigger CI: updated by agent to run full non-DRY_RUN report generation -->

<!-- CI trigger: minimal update to cause GitHub Actions to run full non-DRY_RUN report generation. Agent change on 2025-11-14 -->
