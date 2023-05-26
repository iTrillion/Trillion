import requests
import json
from datetime import datetime, timedelta

# ファイル名の取得
filename = "C:\\TWC\\Log\\" + datetime.now().strftime('%Y%m') + ".txt"

# ファイルを開く
with open(filename, "r") as f:
  data = json.load(f)

# SharePointにデータを書き込む
url = "https://contoso.sharepoint.com/sites/sales/_api/web/lists/getbytitle('Log')/items(1)/set"
headers = {
  "Authorization": "Bearer <your_access_token>",
  "Content-Type": "application/json"
}


payload = json.dumps(data)

print(data)

#response = requests.post(url, headers=headers, data=payload)

#if response.status_code == 200:
print("データの書き込みに成功しました。")
#else:
#  print("データの書き込みに失敗しました。")

