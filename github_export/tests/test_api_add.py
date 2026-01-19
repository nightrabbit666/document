import requests
import json
import os

PROJECT_ID = "5f078857-f623-4612-a601-99105a83e338"
BASE_URL = "http://127.0.0.1:5000"

# Create a dummy image
with open("test_image.png", "wb") as f:
    f.write(os.urandom(1024))

def test_add_entry():
    url = f"{BASE_URL}/api/project/{PROJECT_ID}/entry"
    
    # Login logic
    session = requests.Session()
    try:
        with open("users.json", "r", encoding='utf-8') as f:
            users = json.load(f)
            if not users:
                print("No users found.")
                return
            username = list(users.keys())[0]
            password = users[username]['password']
    except Exception as e:
        print(f"Cannot find user credentials: {e}")
        return

    # Login
    print(f"Logging in as {username}...")
    login_resp = session.post(f"{BASE_URL}/login", data={"username": username, "password": password})
    if login_resp.url.endswith('/login'): # Redirected back to login
         print("Login failed (redirected).")
         return
    
    print("Logged in.")
    
    data = {
        "entry_date": "2023-10-27",
        "sheet_name": "API_TEST",
        "full_title": "API Full Title",
        "deduction_remark": "No deduction"
    }

    files = {
        "removal_photo": ("test_image.png", open("test_image.png", "rb"), "image/png"),
        "manifest_photo": ("test_image.png", open("test_image.png", "rb"), "image/png"),
    }
    
    # Add Entry
    resp = session.post(url, data=data, files=files)
    print(f"Add Entry Status: {resp.status_code}")
    print(f"Response: {resp.text}")

if __name__ == "__main__":
    test_add_entry()
