import json
import os
import shutil
import time

USERS_FILE = 'users.json'
PROJECTS_DIR = os.path.join('work_assistant', 'projects')
SYSTEM_CONFIG_FILE = 'system_config.json'
TOKEN_LOGS_FILE = 'token_logs.json'

DEFAULT_SYSTEM_PROMPT = """你是一個高階文檔自動化架構師。
你的任務：分析 [空白模板] 與 [已填寫範例] 之間的差異，定義出需要填寫的變數參數，並推導出「製表邏輯」。

【輸入資料】
1. 模板類型: {template_type}
2. 結構差異報告 (DeepDiff):
--------------------------------------------------
{formatted_diff_report}
--------------------------------------------------

3. 模板工作表清單: {template_sheets}
4. 已填寫工作表清單: {filled_sheets}
5. 新增的工作表 (動態擴增): {new_sheets}

6. (備用參考) 空白模板結構:
{blank_json}

7. (備用參考) 已填寫結構 (僅若差異報告無效時參考):
{filled_json}

【分析任務】
1. **識別製表邏輯**: 觀察工作表名稱的變化，說明這份報表是如何擴增的？(例如：「依照日期 MMDD 建立新分頁 + 項目名稱」)。
2. **變數定義**: 尋找需要填寫的欄位 (文字、數字、照片)。
3. **圖片偵測**: 若 JSON 中有 `<<IMAGE_PRESENT>>`，務必將其定義為 image 變數。

【輸出 JSON 要求】
請回傳純 JSON 格式，包含：
- logic_summary: (字串) 簡短說明此報表的擴增邏輯與規則 (讓使用者確認它是否正確)。
- date_pattern: (字串, 選填) 偵測到的日期格式規則。
- parameters: (陣列) 變數列表
    - name: 變數英文名稱
    - description: 中文描述 (若為圖片，請標註 "插入照片")
    - type: 資料類型 (string, number, date, list, image)
    - original_text: 對應到的 '原始' 內容 (用於程式定位)
    - example: 對應到的 '填寫' 內容
    - source: 標註來源 (diff_analysis 或 ai_inference)
    - style: (物件) 圖片設定，若 type=image 請務必包含 { "anchor_cell": "偵測到的坐標(如 2,1)", "layout": "smart_center" }

範例輸出:
{{
  "logic_summary": "本報表邏輯為每日新增工作表，命名規則為 '日期 + 項目' (如 1201 廢紙)。",
  "parameters": [
    {{ "name": "...", "type": "image", ... }}
  ]
}}
"""

def load_users():
    if not os.path.exists(USERS_FILE):
        return {}
    with open(USERS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def get_user(username):
    users = load_users()
    return users.get(username)

def verify_user(username, password):
    user = get_user(username)
    if user and user.get('password') == password:
        user['id'] = username
        return user
    return None

def get_all_projects():
    projects = []
    if not os.path.exists(PROJECTS_DIR):
        os.makedirs(PROJECTS_DIR)
        
    for project_id in os.listdir(PROJECTS_DIR):
        config_path = os.path.join(PROJECTS_DIR, project_id, 'config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                try:
                    data = json.load(f)
                    data['id'] = project_id
                    projects.append(data)
                except json.JSONDecodeError:
                    continue
    return projects

def save_project_config(project_id, config_data):
    project_path = os.path.join(PROJECTS_DIR, project_id)
    if not os.path.exists(project_path):
        os.makedirs(project_path)
    
    config_path = os.path.join(project_path, 'config.json')
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, ensure_ascii=False, indent=4)

def get_project_config(project_id):
    config_path = os.path.join(PROJECTS_DIR, project_id, 'config.json')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None

def get_project_entries(project_id):
    path = os.path.join(PROJECTS_DIR, project_id, 'entries.json')
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return []
    return []

def save_project_entry(project_id, entry_data):
    path = os.path.join(PROJECTS_DIR, project_id, 'entries.json')
    entries = []
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            try:
                entries = json.load(f)
            except:
                pass
    
    entries.append(entry_data)
    
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(entries, f, ensure_ascii=False, indent=4)

def delete_project_entry(project_id, entry_id):
    path = os.path.join(PROJECTS_DIR, project_id, 'entries.json')
    if not os.path.exists(path): return
    
    entries = []
    with open(path, 'r', encoding='utf-8') as f:
        try:
            entries = json.load(f)
        except:
            return
            
    new_entries = [e for e in entries if e.get('id') != entry_id]
    
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(new_entries, f, ensure_ascii=False, indent=4)

def get_system_config():
    """Load system configuration, creating default if not exists."""
    if not os.path.exists(SYSTEM_CONFIG_FILE):
        default_config = {
            "ai_prompt_template": DEFAULT_SYSTEM_PROMPT,
            "model_name": "gemini-3-flash-preview",
            "ui_settings": {
                "theme": "light",
            }
        }
        save_system_config(default_config)
        return default_config
    
    with open(SYSTEM_CONFIG_FILE, 'r', encoding='utf-8') as f:
        try:
            config = json.load(f)
            if "ai_prompt_template" not in config:
                config["ai_prompt_template"] = DEFAULT_SYSTEM_PROMPT
            return config
        except:
             return {"ai_prompt_template": DEFAULT_SYSTEM_PROMPT}

def save_system_config(config):
    with open(SYSTEM_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def log_token_usage(project_id, tokens):
    logs = []
    if os.path.exists(TOKEN_LOGS_FILE):
        with open(TOKEN_LOGS_FILE, 'r', encoding='utf-8') as f:
            try:
                logs = json.load(f)
            except:
                pass
    
    entry = {
        "timestamp": time.time(),
        "date": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
        "project_id": project_id,
        "tokens": tokens
    }
    logs.append(entry)
    if len(logs) > 1000:
        logs = logs[-1000:]
        
    with open(TOKEN_LOGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(logs, f, ensure_ascii=False, indent=4)

def get_token_usage_stats():
    logs = []
    if os.path.exists(TOKEN_LOGS_FILE):
        with open(TOKEN_LOGS_FILE, 'r', encoding='utf-8') as f:
            try:
                logs = json.load(f)
            except:
                pass
    return logs

def delete_project(project_id):
    path = os.path.join(PROJECTS_DIR, project_id)
    if os.path.exists(path):
        try:
            shutil.rmtree(path)
            return True
        except Exception as e:
            print(f"Error deleting project: {e}")
            return False
    return False
