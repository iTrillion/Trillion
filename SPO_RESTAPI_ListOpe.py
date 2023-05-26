import json
import requests
from datetime import datetime, timedelta
#####################################################################
### 基本情報 ###
target_SiteName = 'msteams_2d68ac'
client_id = '2ec6495c-485e-4be1-9b1e-5c8fa1fdf461@d1c1335e-f582-42a9-b6fe-5e1a16eb9bc8'
client_secret = 'LW4HvfDKjwXdq+lGxxIpwrMaloGOaCuYHyJ89cS1Y3o='
target_ListName = 'PPTM_Summery'
target_ListEntityName = 'SP.Data.PPTM_x005f_SummeryListItem'
#####################################################################


#####################################################################
### アクセストークン取得 ###
url = 'https://accounts.accesscontrol.windows.net/d1c1335e-f582-42a9-b6fe-5e1a16eb9bc8/tokens/OAuth/2'
data = {
    'grant_type':'client_credentials',
    'client_id':client_id,
    'client_secret':client_secret,
    'resource':'00000003-0000-0ff1-ce00-000000000000/toyotajp.sharepoint.com@d1c1335e-f582-42a9-b6fe-5e1a16eb9bc8',
}
headers = {
    'Content-Type':'application/x-www-form-urlencoded',
}

r = requests.post(url, data=data, headers=headers)
json_token = json.loads(r.text)
#####################################################################
### 入力データ作成 ###
input_data = {
    'Title': datetime.strftime(datetime.now() +timedelta(hours=9),"%Y/%m/%d %H:%M:%S"),
    'Page1': "001",
    'Page2': "002",
    'Page3': "003",
    'Page4': "004",
    'Page5': "005",
    'Page6': "006",
    'Page7': "007",
    'Page8': "008",
    'Page9': "009",
    'Page10': "010",
    'Page11': "011",
    'Page12': "012",
    'Page13': "013",
    'Page14': "014",
    'Menu1': "101",
    'Menu2': "102",
    'Menu3': "103",
    'Menu4': "104",
}


FlName =  now.strftime('%Y%m%') +'.log'

# ファイルからテキストデータを読み込む
with open('C:\TWC\Log\'+Flname, "r") as f:
  text = f.read()

# テキストデータをJSON形式に変換する
json_data = json.loads(text)

response = requests.post(url, headers=headers, data=json.dumps(json_data))

#####################################################################
### リストアイテム作成 ###
url = "https://toyotajp.sharepoint.com/sites/"+ target_SiteName +"/_api/web/lists/GetByTitle('"+ target_ListName +"')/items"
data = '''{
    "__metadata": {
        "type": "'''+ target_ListEntityName +'''"
    },
    "Title": "''' + input_data['Title'] + '''",
    "Page1": "''' + input_data['Page1'] + '''",
    "Page2": "''' + input_data['Page2'] + '''",
    "Page3": "''' + input_data['Page3'] + '''",
    "Page4": "''' + input_data['Page4'] + '''",
    "Page5": "''' + input_data['Page5'] + '''",
    "Page6": "''' + input_data['Page6'] + '''",
    "Page7": "''' + input_data['Page7'] + '''",
    "Page8": "''' + input_data['Page8'] + '''",
    "Page9": "''' + input_data['Page9'] + '''",
    "Page10": "''' + input_data['Page10'] + '''",
    "Page11": "''' + input_data['Page11'] + '''",
    "Page12": "''' + input_data['Page12'] + '''",
    "Page13": "''' + input_data['Page13'] + '''",
    "Page14": "''' + input_data['Page14'] + '''",
    "Menu1": "''' + input_data['Menu1'] + '''",
    "Menu2": "''' + input_data['Menu2'] + '''",
    "Menu3": "''' + input_data['Menu3'] + '''",
    "Menu4": "''' + input_data['Menu4'] + '''"
}'''
headers = {
    'Authorization': 'Bearer ' + json_token['access_token'],
    'Accept':'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose',
    'Content-Length': str(len(data)),
}

print(data)
l = requests.post(url, data=data, headers=headers)
print('---------リストアイテム作成---------')
# レスポンスを確認する
if response.status_code == 200:
  print("POST succeeded")
else:
  print("POST failed")


print(l.text)
#####################################################################

