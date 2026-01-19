
def get_system_config():
    """Load system configuration, creating default if not exists."""
    if not os.path.exists(SYSTEM_CONFIG_FILE):
        default_config = {
            "ai_prompt_template": DEFAULT_SYSTEM_PROMPT,
            "model_name": "gemini-3-flash-preview",
            "ui_settings": {
                "theme": "light",
                "flow_chart_steps": [
                    {"step": 1, "label": "上傳模板", "desc": "上傳 Word/Excel 空白模板與參考範例"},
                    {"step": 2, "label": "AI 分析", "desc": "系統自動分析差異與邏輯"},
                    {"step": 3, "label": "專案建立", "desc": "確認參數並建立專案"},
                    {"step": 4, "label": "日常填寫", "desc": "使用者填寫表單並產出文件"}
                ]
            }
        }
        save_system_config(default_config)
        return default_config
    
    with open(SYSTEM_CONFIG_FILE, 'r', encoding='utf-8') as f:
        try:
            config = json.load(f)
             # Ensure defaults exist
            if "ai_prompt_template" not in config:
                config["ai_prompt_template"] = DEFAULT_SYSTEM_PROMPT
            return config
        except:
             return {"ai_prompt_template": DEFAULT_SYSTEM_PROMPT}

def save_system_config(config):
    with open(SYSTEM_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def log_token_usage(project_id, tokens):
    """
    tokens = {'prompt_tokens': 100, 'candidates_tokens': 50, 'total_tokens': 150}
    """
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
    
    # Keep last 1000 entries
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
