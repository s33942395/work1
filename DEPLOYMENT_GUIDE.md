# 🚀 Streamlit Cloud 部署指南

## 準備工作 ✅

已完成以下準備：
- ✅ 更新 `requirements.txt`（包含所有必要套件）
- ✅ 創建 `.streamlit/config.toml`（優化配置）
- ✅ 創建 `README.md`（專案說明）
- ✅ 創建 `.gitignore`（排除不必要檔案）
- ✅ 推送到 GitHub（repository: s33942395/work1）

## 部署步驟

### 1. 前往 Streamlit Cloud
訪問：https://share.streamlit.io/

### 2. 登入
使用您的 GitHub 帳號登入（s33942395）

### 3. 創建新應用
1. 點擊右上角的 **"New app"** 按鈕
2. 填寫以下資訊：
   - **Repository**: `s33942395/work1`
   - **Branch**: `main`
   - **Main file path**: `cloud_app.py`
   - **App URL** (optional): 自訂網址，例如 `governance-survey-analysis`

### 4. 高級設置（Advanced settings）
點擊 "Advanced settings" 可以設定：
- **Python version**: 選擇 3.11 或更新版本
- **Secrets**: 如果需要加密資料，可以在這裡設定

### 5. 部署
點擊 **"Deploy!"** 按鈕

### 6. 等待部署完成
- 初次部署約需 2-5 分鐘
- 您可以在日誌中看到安裝套件的進度
- 部署完成後會自動開啟應用程式

## 部署後的應用網址

部署完成後，您的應用程式將可通過以下網址訪問：
```
https://share.streamlit.io/s33942395/work1/main/cloud_app.py
```

或者如果您自訂了網址：
```
https://<your-custom-url>.streamlit.app
```

## 🔄 更新應用程式

當您更新程式碼並推送到 GitHub 後：

1. **自動部署**：Streamlit Cloud 會自動偵測變更並重新部署
2. **手動重啟**：如果需要，可以在應用管理頁面點擊 "Reboot" 按鈕

## 📊 管理應用程式

在 Streamlit Cloud 管理界面，您可以：
- 查看應用狀態和使用量
- 查看日誌（Logs）
- 重啟應用（Reboot）
- 刪除應用（Delete）
- 設定 Secrets（如需要）

## 💡 注意事項

### 檔案上傳
- 在雲端環境中，使用者需要每次上傳 CSV 檔案
- 檔案大小限制：200MB（已在 config.toml 中設定）

### 資料安全
- 上傳的檔案不會永久儲存
- 每次 session 重新整理後需要重新上傳
- 如需要持久化資料，考慮使用資料庫或雲端儲存

### 效能優化
- 使用 `@st.cache_data` 快取資料處理結果
- 大型資料集可能需要較長載入時間

## 🐛 故障排除

### 部署失敗
1. 檢查 `requirements.txt` 是否正確
2. 查看部署日誌（Logs）找出錯誤
3. 確認 Python 版本相容性

### 應用程式錯誤
1. 在 Streamlit Cloud 查看 "Logs"
2. 確認所有必要套件都已安裝
3. 檢查檔案路徑是否正確

### 記憶體不足
- Streamlit Cloud 免費版有記憶體限制
- 考慮優化資料處理邏輯
- 或升級到付費方案

## 🔗 相關連結

- Streamlit Cloud: https://share.streamlit.io/
- Streamlit 文檔: https://docs.streamlit.io/
- GitHub Repository: https://github.com/s33942395/work1

## ✨ 下一步

部署完成後，您可以：
1. 分享應用網址給團隊成員
2. 測試所有功能是否正常運作
3. 根據需求進一步優化使用者體驗

祝部署順利！🎉
