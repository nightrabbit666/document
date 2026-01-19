# 行政文書助手 (Document Assistant) - 下一代系統建置規格書

本規格書詳細定義了「行政文書助手」系統的架構、邏輯、UI/UX 設計及開發注意事項。請依照此文件在乾淨的環境中重建系統。

---

## 1. 系統架構總覽 (System Architecture)

本系統為一個基於 Web 介面的自動化文档處理平台，利用 Generative AI (Google Gemini) 解析非結構化文件（Word/Excel），識別動態參數，並建立標準化模板以供後續快速生成報告。

### 技術堆疊 (Tech Stack)
*   **後端框架**: Python Flask (輕量級 Web Server)
*   **前端技術**: HTML5, JavaScript (原生), Tailwind CSS (CDN版本)
*   **AI 核心**: Google Gemini 3 Flash (透過 `google-generativeai` 庫)
*   **文件處理**:
    *   `python-docx`: 讀取與寫入 Word 內容
    *   `docxtpl`: Jinja2 風格的 Word 模板替換引擎
    *   `openpyxl` / `pandas`: 處理 Excel 數據導入
*   **資料存儲**: JSON 文件系統 (輕量化，無需 SQL 資料庫)

---

## 2. 目錄結構規範 (Directory Structure)

重建系統時，請嚴格遵守此結構以確保路徑引用正確。

```text
Project_Root/
├── .venv/                      # Python 虛擬環境
├── requirements.txt            # 依賴清單
├── run_server.ps1              # 啟動腳本 (PowerShell)
├── users.json                  # 用戶認證資料 (模擬資料庫)
└── work_assistant/             # 核心應用目錄
    ├── txtapp.py               # [核心] 主程式入口 & 後端邏輯
    ├── database.py             # (選用) 封裝 JSON 讀寫邏輯
    ├── uploads/                # 臨時上傳區 (需定期清理)
    ├── projects/               # 專案設定檔存儲區
    │   └── {Project_ID}/       # 每個專案一個資料夾
    │       └── config.json     # 專案參數配置
    ├── static/
    │   ├── js/                 # 前端複雜邏輯 (可選)
    │   └── css/
    └── templates/              # HTML 視圖
        ├── index.html          # 首頁 (儀表板)
        ├── login.html          # 登入頁
        ├── project_setup.html  # [核心] 新增專案精靈 (Step 1-3)
        └── project_form.html   # 填寫報表頁
```

---

## 3. 核心環境依賴 (requirements.txt)

```text
flask
flask-login
google-generativeai
python-docx
docxtpl
pandas
openpyxl
werkzeug
```

---

## 4. 後端核心邏輯設計 (Backend Logic)

### 4.1 檔案上傳與安全處理 (File Handling)
**邏輯說明**：
1.  **檔名安全化**：原始中文檔名在某些 OS 會亂碼，必須使用 UUID 重命名存檔，但在資料庫/前端回傳時需保留原始檔名以供辨識。
2.  **允許格式**：除了 `.docx` (模板)，必須支援 `.xlsx/.xls` (資料源) 以進行交叉比對。

**Python 實作範例**：
```python
def upload_handler():
    # 產生安全檔名，保留副檔名
    file_ext = os.path.splitext(file.filename)[1]
    safe_filename = f"{uuid.uuid4()}{file_ext}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
    file.save(filepath)
    # ...後續邏輯
```

### 4.2 AI 分析引擎邏輯 (The Brain)
這是系統的核心，建議封裝於函數中 (如 `analyze_template_with_gemini`)。

**流程**：
1.  **提取文字**：使用 `python-docx` 讀取 Word 的所有 Paragraphs 和 Tables。
2.  **提取上下文**：如果有上傳 Excel 或 已填寫的 Word，同時提取其內容作為「參考資料 (Context)」。
3.  **Prompt 建構**：將「空白範本內容」+「參考資料」發送給 Gemini。
4.  **AI 任務**：
    *   找出文件中的「變數」（例如：日期、金額、專案名稱）。
    *   如果能在「參考資料」中找到對應值，則自動提取作為「範例值 (Example)」。
    *   推斷變數類型 (string, number, date, list)。
5.  **JSON 解析**：Gemini 回傳 JSON 格式的參數列表。

### 4.3 文檔生成引擎 (Document Generation)
使用 `docxtpl` 庫進行模板渲染。

**邏輯**：
1.  使用者在前端確認參數名稱（例如 `{{report_date}}`）。
2.  後端將使用者上傳的 Word 轉換為 Jinja2 模板（將由 AI 識別出的文字替換為 `{{變數名}}`）。
3.  **防呆機制**：必須精確比對文字，避免破壞 Word 的 XML 結構。

---

## 5. 前端 UI/UX 設計規範

### 5.1 色彩系統 (Color Palette)
使用 Tailwind CSS 類別：
*   **主色調 (Primary)**: `emerald-600` (綠色系，代表環保/文件/通過)
*   **輔助色 - Excel**: `blue-600` (藍色系，代表數據)
*   **輔助色 - 參考檔**: `purple-600` (紫色系，代表智慧/比對)
*   **背景**: `slate-50` (淺灰)
*   **文字**: `slate-800` (深灰)

### 5.2 互動流程設計 (User Flow)

#### 頁面：新增專案 (Project Setup)
採 **Step-by-Step Wizard** 設計。

**Step 1: 上傳與分析 (Upload & Analyze)**
*   **Drag & Drop Zones (拖放區)**：
    *   **Zone A (必要)**: 空白 Word 範本。
    *   **Zone B (選用)**: 已填寫的舊 Word 檔 (紫色)。(新增功能：用於AI學習舊資料)
    *   **Zone C (選用)**: Excel 資料檔 (藍色)。(新增功能：用於AI對應欄位)
*   **互動邏輯**：每個區域需獨立處理 `drop` 和 `click` 事件，避免 ID 衝突。
*   **回饋**：上傳後必須顯示檔名與檔案大小，並隱藏上傳提示圖標。
*   **Loading**：按下分析後，顯示進度條 (Progress Bar)，輪詢後端狀態。

**Step 2: 參數確認 (Verify)**
*   左側：列出 AI 識別到的參數 (卡片式設計)。
*   右側：AI Chatbot (允許使用者用自然語言修改參數，如「把日期格式改成民國年」)。
*   **保存**：確認無誤後，將配置寫入 `config.json`。

**Step 3: 完成 (Finish)**
*   顯示專案摘要，提供「立即使用」或「下載模板」按鈕。

### 5.3 關鍵 UI 程式碼邏輯 (JavaScript)
*   **事件監聽**：HTML 標籤內**禁止**寫 `onclick`，統一使用 `addEventListener` 綁定 DOM 元素。
*   **狀態管理**：使用 `sessionId` (全域變數) 識別當前分析任務。
*   **Polling (輪詢)**：使用 `setInterval` 每秒檢查後端分析進度，直到 `status === 'completed'`，並加上逾時處理。

---

## 6. AI Prompt 工程 (核心指示詞)

在 `txtapp.py` 中發送給 Gemini 的 Prompt 結構應如下：

```text
你是一個專業的文件分析助手。我會提供你一份 Word 文件的內容結構。
你的任務是識別出這份文件中「應該是動態填寫」的部分（變數）。

【輸入資料】
1. 文件文本內容: {content}
2. (選用) 參考Excel數據: {excel_data}
3. (選用) 參考舊文件內容: {filled_content}

【輸出要求】
請回傳純 JSON 格式，包含一個 'parameters' 陣列，每個物件包含：
- name: 建議的變數英文名稱 (如 report_date, total_amount)
- description: 變數的中文描述 (如 報告日期, 總金額)
- type: 資料類型 (string, number, date, list)
- context: 該變數在文中出現的上下文句子
- example: 從參考資料中找到的範例值 (如果有的話)

【重要規則】
- 如果有提供Excel資料，請嘗試將Excel的欄位名稱與文件內容進行匹配。
- 忽略頁碼、頁眉、頁腳中的固定文字。
```

---

## 7. 逐步建置與修復檢查表 (Checklist)

在重建過程中，請依序檢查以下重點，這些都是曾發生過的 Bug 點：

### 環境與後端
- [ ] **Python 版本**: 確認使用 Python 3.10+。
- [ ] **API Key**: `GEMINI_API_KEY` 是否已設定並有效。
- [ ] **允許格式**: `ALLOWED_EXTENSIONS` 是否包含 `{'xlsx', 'xls', 'docx', 'doc'}`。
- [ ] **檔名處理**: 是否使用 `uuid` 避免中文亂碼問題。
- [ ] **Request 處理**: 是否正確接收 `request.files['filled_template']` 與 `request.files['excel']`。

### 前端 UI
- [ ] **ID 唯一性**: 檢查 `project_setup.html` 是否有重複的 `id="fileInfo"` (應只有一個或使用不同 ID)。
- [ ] **事件衝突**: 上傳按鈕 (Button) 是否**移除**了 `onclick` 屬性，完全交由外層 `dropZone` 的監聽器處理 (避免觸發兩次)。
- [ ] **區塊顯示**: 確認 `Step 2` 容器預設為 `hidden`。
- [ ] **錯誤回饋**: 當 API 回傳非 JSON (如 500 Error HTML) 時，JS 是否有 `try-catch` 處理?

### 流程測試
- [ ] **Excel 上傳**: 確認上傳 `.xlsx` 不會報錯 "Not a Word file"。
- [ ] **空白範本**: 確認點擊上傳區能正常彈出檔案選擇視窗。
- [ ] **分析結果**: 確認 AI 分析完畢後，參數列表能正確渲染在畫面上。
