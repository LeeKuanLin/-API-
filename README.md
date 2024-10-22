import requests
import pandas as pd

# 第一階段：獲取會話 ID
endpoint_login = "https://slt.eup.tw:8443/"
getSessionIdUrl = "Eup_Login_SOAP/Eup_Login_SOAP"

login_data = {
    'Param': '{"MethodName":"Login","CoName":"43622035","Account":"43622035","Password":"0000"}'
}

response_login = requests.post(endpoint_login + getSessionIdUrl, data=login_data)
session_info = response_login.json()

# 打印完整的伺服器響應
#print("Login Response:", session_info)
session_id = session_info["SESSION_ID"]


getSessionIdUrl = "Eup_Statistics_SOAP/Eup_Statistics_SOAP"

formData = {
    'Param': '{"Cust_IMID":"7867","custImid":"7867","Cust_ID":"5028213","custId":"5028213","Team_ID":"5027043","teamId":"5027043","SESSION_ID":"' + session_id +'","StartDate":"2024-10-15 00:00:00","EndDate":"2024-10-21 23:59:59","Driver_ID":null,"IsSelf":true,"MethodName":"GetCarRecordByDriver"}'
}

response = requests.post(endpoint_login+getSessionIdUrl, data = formData)


import pandas as pd

# 假设你的 data 是从 API 获得的数据
data = response.json()["result"]
df = pd.DataFrame(data)
columns_mapping = {
    'Car_Driver': '駕駛姓名',
    'Car_Number': '車牌',
    'StartTime': '開始時間',
    'EndTime': '結束時間',
    'StartAddress': '開始地址',
    'EndAddress': '結束地址',
    'Account': '帳號',
    
}
new_column_order = list(columns_mapping.keys())
df = df[new_column_order]
df.rename(columns=columns_mapping, inplace=True)
# 刪除“駕駛姓名”中的“台中”、“高雄”、“台北”，並創建新列“駕駛姓名（無特定字眼）”
remove_words = ['台中', '高雄', '台北','桃園','新竹','台南','連昌']
pattern = '|'.join(remove_words)
df['駕駛姓名（無特定字眼）'] = df['駕駛姓名'].str.replace(pattern, '', regex=True)
df['駕駛姓名（無特定字眼）'] = df['駕駛姓名（無特定字眼）'].str.lstrip('-')

# 調整列的順序，把“駕駛姓名（無特定字眼）”放到第二列
column_order = ['駕駛姓名', '駕駛姓名（無特定字眼）', '車牌', '開始時間', '結束時間', '開始地址', '結束地址', '帳號']
df = df[column_order]


# 指定保存文件的路径，使用原始字符串来避免转义字符问题
file_path = r"C:\Users\eup-f\Downloads\python打卡紀錄\car_record_by_driver.xlsx"

try:
    # 将 DataFrame 导出到指定路径的 Excel 文件
    df.to_excel(file_path, index=False)
    print(f"Data successfully exported to {file_path}")
except Exception as e:
    print(f"An error occurred: {e}")
