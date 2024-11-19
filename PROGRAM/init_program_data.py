import os
from dotenv import load_dotenv
load_dotenv("MASTER_DATA/.env")
import datetime
import pandas as pd
from google.oauth2.service_account import Credentials
import gspread

def make_dir(year_month_date):
    os.makedirs(f"TRAN_DATA//{year_month_date}", exist_ok=True)

    with open(f"TRAN_DATA//{year_month_date}//subject.txt", mode="w", encoding="utf-8") as f:
        f.write("このテキストを削除して、件名を入力してください")
    
    with open(f"TRAN_DATA//{year_month_date}//parts_name.txt", mode="w", encoding="utf-8") as f:
        f.write("このテキストを削除して、部品名を入力してください\n改行したい場合は「<br>」と入力してください")

def str_to_list(str):
    return str.replace("[", "").replace("]", "").replace("'", "").split(", ")

def get_spread_sheet(spreadsheet_key):
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    credentials = Credentials.from_service_account_file(
        os.environ.get("AUTH_PATH"),
        scopes=scopes
    )

    gc = gspread.authorize(credentials)
    
    spreadsheet = gc.open_by_key(spreadsheet_key)

    return spreadsheet

def spreadsheet_to_df(spreadsheet, sheet_name):
    df = pd.DataFrame(spreadsheet.worksheet(sheet_name))
    df.columns = df.iloc[0]
    df = df.drop(0, axis=0)
    return df

def make_mail_status_csv(year_month_date):
    spreadsheet = get_spread_sheet(os.environ.get("SPRADSHEET_KEY"))
    df = spreadsheet_to_df(spreadsheet, os.environ.get("SHEET_NAME"))

    df = df[str_to_list(os.environ.get("USE_COLUMNS"))]
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

    df.to_csv(f".//TRAN_DATA//{year_month_date}//mail_status.csv", index=False)

if __name__ == "__main__":
    # year-month-dataを作成
    year_month_date = datetime.datetime.now().strftime("%Y%m%d")
    
    make_dir(year_month_date)
    make_mail_status_csv(year_month_date)
