import requests

url = "https://excelchatbot-n8n-production.up.railway.app/webhook/task-manager"

payload = {
    "action": "test"
}

r = requests.post(url, json=payload)

print("Status:", r.status_code)
print("Response:", r.text)