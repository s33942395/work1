# 圖表生成功能修復說明

## 問題描述
Word 報告中出現「圖表生成失敗」的錯誤訊息。

## 根本原因
Kaleido（Plotly 的靜態圖像生成引擎）需要 Chrome/Chromium 瀏覽器及其系統依賴套件才能運作。

## 解決方案

### 1. 安裝系統依賴套件
```bash
sudo apt-get update && sudo apt-get install -y \
  libnss3 \
  libatk-bridge2.0-0 \
  libcups2 \
  libxcomposite1 \
  libxdamage1 \
  libxfixes3 \
  libxrandr2 \
  libgbm1 \
  libxkbcommon0 \
  libpango-1.0-0 \
  libcairo2 \
  libasound2
```

### 2. 安裝並配置 Kaleido
```bash
pip install -U kaleido
python -c "import kaleido; kaleido.get_chrome_sync()"
```

### 3. 測試圖表生成
```python
import plotly.graph_objects as go

fig = go.Figure(data=[go.Bar(x=['A', 'B', 'C'], y=[1, 2, 3])])
fig.write_image('/tmp/test.png', width=800, height=600)
print("✓ 圖表生成成功")
```

## 驗證結果
- ✅ 圖表可正常生成 PNG 檔案
- ✅ Word 報告包含圖表（每個議題 2 張圖）
- ✅ 檔案大小正常（>100KB）
- ✅ 圖表品質良好（1000x600，scale=2）

## 技術細節
- **圖表引擎**: Plotly + Kaleido
- **瀏覽器**: Chromium（透過 Kaleido 管理）
- **圖表格式**: PNG（1000x600 px，scale=2）
- **嵌入方式**: python-docx 的 `add_picture()` 方法

## 圖表類型
每個議題包含 2 個圖表：
1. **公司方 vs 投資方比較圖**：群組長條圖
2. **階段比較圖**（若有階段資料）：多階段群組長條圖

## 部署建議
在新環境部署時，請確保執行上述系統依賴套件的安裝命令。這些套件通常在完整的 Linux 系統中已安裝，但在 Docker 容器或精簡版 OS 中可能缺少。

---
修復日期：2025-11-12  
修復版本：Kaleido 1.x with Chrome support
