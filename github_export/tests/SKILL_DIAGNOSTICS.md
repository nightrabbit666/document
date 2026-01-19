# AI 診斷與測試技能庫 (Diagnostic Skills)

本文件記錄了系統重構開發過程中的關鍵測試程序與除錯技巧，方便日後開發者進行功能驗證與問題排除。

## 1. 原理概述

本系統採用「混合式驗證」策略：
- **DeepDiff**: 用於 Python 原生結構比對 (速度快，適合表格文字)。
- **Gemini AI**: 用於語意解讀與補強 (當 DeepDiff 失效或針對圖片時)。

---

## 2. 測試腳本說明

這些腳本位於 `tests/` 目錄下。

### 2.1 檢查 AI 模型可用性 (`check_models.py`)

**用途**：當系統回報 404 Model Not Found 時，使用此腳本列出目前 API Key 可存取的所有模型列表。

**如何執行**:
```bash
python tests/check_models.py
```

**輸出判讀**:
- 若列表為空：API Key 無效或權限不足。
- 若有出現模型：請將系統中的 `txtapp.py` 模型名稱更新為列表中最新且支援 `generateContent` 的版本 (如 `gemini-3-flash-preview` 或 `gemini-1.5-pro`)。

---

### 2.2 重現解析問題 (`reproduce_issue.py`)

**用途**：脫離 Flask 網站環境，直接對 Excel/Word 進行 "原子測試"。這能幫助我們釐清「是程式讀不到檔案」還是「AI 判斷錯誤」。

**如何執行**:
```bash
python tests/reproduce_issue.py
```

**適用場景**:
- **圖片讀取失敗**: 當 `openpyxl` 回報 "0 images" 但 Excel 裡明明有圖時。
- **比對無效**: 當網頁顯示 "未偵測到變數" 時，跑這個腳本查看底層 DeepDiff 的 raw output。

**腳本邏輯**:
1. 硬編碼讀取 `參考/` 資料夾中的 `空白.xlsx` 與範例檔。
2. 嘗試列印出所有 Sheet 的圖片數量與坐標。
3. 嘗試發送簡單 Prompt 給 AI。

---

### 2.3 檢查 Excel 圖片路徑與狀態 (`check_excel_images.py`)

**用途**：診斷為何 Excel 輸出沒有圖片。此腳本專門測試 `Pillow` 庫的安裝狀態以及圖片插入邏輯。

**關鍵發現 (2026-01-19 Lesson Learned)**:
- **Pillow 衝突**: `openpyxl` 依賴 `Pillow` 處理圖片，若環境未正確安裝或有多個衝突版本，圖片不會報錯但會直接消失。
- **路徑解析**: `openpyxl.drawing.image.Image` 接受絕對路徑。若使用相對路徑，必須確保 CWD (Current Working Directory) 正確。

---

## 3. 常見問題排除 (Troubleshooting)

### Q1: Excel 圖片讀取不到 (0 Images)
- **原因**: 圖片可能是 `Floating Object` (浮動) 或是 `Grouped` (群組)，`openpyxl` 只能讀取錨定好的圖片。
- **解法**: 
    1. 嘗試將圖片「置於儲存格內」(Right Click -> Place in Cell)。
    2. 若程式無法解決，信賴 AI 的 Fallback 機制 (AI 直接看儲存格文字上下文)。

### Q2: AI 回應 404
- **解法**: 執行 `check_models.py`，確認 `google-generativeai` 套件版本與 Google 雲端開放的模型是否一致。

### Q3: 0 Byte Upload
- **解法**: 這通常是前端顯示問題。後端 `save_uploaded_file` 其實已經成功儲存。檢查 `api_upload` 回傳的 JSON 是否包含 `size` 欄位。

### Q4: 圖片上傳失敗 (後端回傳 None)
- **原因**: Flask 的 `ALLOWED_EXTENSIONS` 設定過嚴，預設只開啟 `docx`, `xlsx`。若上傳 `jpg`, `png`，`allowed_file` 函式會回傳 `False`，導致檔案被靜默忽略。
- **解法**: 確保 `txtapp.py` 中 `ALLOWED_EXTENSIONS` 包含 `{'png', 'jpg', 'jpeg'}`。

---

## 4. 進階邏輯分析技能 (Advanced Logic Patterns)

### 4.1 動態報表偵測 (Dynamic Reporting)
針對每日月報表 (Monthly Report with Daily Entries) 類型的文件，我們發現以下特徵：

1.  **結構特徵**:
    - 模板檔通常只包含一個空白 Sheet 或範例 Sheet。
    - 實作檔 (Filled) 會根據日期大量擴增 Sheet。
    
2.  **偵測技術**:
    - 在 Python 端 `textapp.py` 需提取 `wb.sheetnames`。
    - 計算 `new_sheets = filled_sheets - template_sheets`，這群新增的 Sheet 就蘊含了製表邏輯。
    
3.  **Prompt 策略**:
    - 必須明確告知 AI：「這是一份會隨日期擴增的文件」。
    - 將 `new_sheets` 列表餵給 AI，請求它分析 naming convention (如 `1201 廢紙` -> `MMDD + Category`).

### 4.2 圖片智能排版 (Smart Image Layout)
針對 Excel 圖片插入的精確定位問題，我們發展出以下邏輯：

1.  **智慧錨定 (Smart Anchor)**:
    - 由於 Excel 圖片是懸浮層 (Drawing Layer)，若插入點 (Anchor) 的列高太小 (如標題列)，圖片會遮擋文字。
    - **邏輯**: 偵測 Anchor 所在列的 `Height`。若高度 < 60pt，自動將 Anchor 下移一列 (Body Cell)。

2.  **合併儲存格置中 (Merged Cell Centering)**:
    - `openpyxl` 預設圖片置於左上角。
    - **邏輯**:
        - 偵測 Anchor 是否屬於 `merged_cells`。
        - 計算合併區塊的總像素寬度與高度。
        - 計算圖片的 EMU (English Metric Unit)，並設定 `OneCellAnchor` 的 `colDown` (Offset) 值，使 `(CellWidth - ImageWidth) / 2` 成為偏移量，達成水平/垂直置中。

### 4.3 Token 監控
- Google Gemini API 回傳的 `usage_metadata` 包含 `prompt_token_count` 與 `candidates_token_count`。
- 監控此數據可幫助預估成本。目前的分析約消耗 1k~3k Tokens/次。

---

## 5. 維護記錄
- **2025-01-16**: 
    - 建立測試庫。
    - 修正 AI 模型為 `gemini-3-flash-preview`。
    - 發現 `openpyxl` 對使用者提供的 Excel 圖片支援度有限，強化 AI 介入邏輯。
    - 新增「製表邏輯分析」，支援動態 Sheet 擴增偵測。
- **2026-01-19**:
    - **圖片上傳修復**: 擴增 `ALLOWED_EXTENSIONS` 支援圖片格式。
    - **圖片排版引擎**: 實作 `Smart Anchor` 與 `Centering` 演算法，解決圖片跑版問題。
    - **依賴修復**: 解決 PowerShell 導致 `requirements.txt` 編碼錯亂 (Null Bytes) 問題。
