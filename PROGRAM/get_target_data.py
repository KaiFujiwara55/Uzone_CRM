import json
import pandas as pd
from google.oauth2.service_account import Credentials
import gspread

with open("C:\\Users\\info\\Desktop\\Uzone_CRM\\MASTER_DATA\\ENV_PATH.json") as f:
    ENV_PATH = json.load(f)

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

credentials = Credentials.from_service_account_file(
    ENV_PATH["auth_path"],
    scopes=scopes
)

gc = gspread.authorize(credentials)

spreadsheet_key = "1b4hvRc4BVCqhdsHIEniXkmfwolU9moyvCqDps1eb9ko"

spreadsheet = gc.open_by_key(spreadsheet_key)
df = pd.DataFrame(spreadsheet.worksheet("送信対象").get_all_values())
df.columns = df.iloc[0]
df = df.drop(0, axis=0)
df = df[["SF登録名(UZone登録名)", "メールアドレス"]]
df.columns = ["account_name", "mail_address"]

tmp_df = df[df["mail_address"].str.contains("、")]
df = df[~df["mail_address"].str.contains("、")]
for idx, row in tmp_df.iterrows():
    mail_address_list = row["mail_address"].split("、")
    for mail_address in mail_address_list:
        new_row = pd.Series([row["account_name"], mail_address], index=["account_name", "mail_address"])
        df = pd.concat([df, pd.DataFrame([new_row])])

df = df.reset_index(drop=True)
for idx, row in df.iterrows():
    account_name = row["account_name"]
    if account_name[-1] == "店":
        df.iloc[idx, 0] = df.iloc[idx, 0][:-1]

df["is_send"] = "未"

df.to_csv(".//TRAN_DATA//20241104//mail_status.csv", index=False)
