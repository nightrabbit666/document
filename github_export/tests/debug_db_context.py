import requests
import json
import os
import sys

# Add project root to path
sys.path.append(r"C:\Users\raven.yeh\Desktop\專案拓展部\09_AI增能\行政助手(vscode重建)")
from work_assistant.txtapp import app, database

PROJECT_ID = "5f078857-f623-4612-a601-99105a83e338"

print(f"Checking project: {PROJECT_ID}")

# Check config first
config = database.get_project_config(PROJECT_ID)
print(f"Loaded Config: {len(config.get('parameters', []))} params")
for p in config.get('parameters', []):
    print(f" - {p['name']} ({p['type']})")

if 'removal_photo' in [p['name'] for p in config['parameters']]:
    print("Config HAS removal_photo.")
else:
    print("Config MISSING removal_photo.")

# Check entries
entries = database.get_project_entries(PROJECT_ID)
print(f"Entries: {len(entries)}")
if entries:
    # Print the last 2 entries to see history
    for i in range(min(2, len(entries))):
        entry = entries[-(i+1)]
        print(f"Entry {-(i+1)} ({entry['date']}): Keys={list(entry['data'].keys())}")
        print(f"   Data: {entry['data']}")

