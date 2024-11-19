import os
import json
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

def spreadsheet_to_df(spreadsheet_key, sheet_name):
    spreadsheet = get_spread_sheet(spreadsheet_key)
    df = pd.DataFrame(spreadsheet.worksheet(sheet_name).get_all_values())
    df.columns = df.iloc[0]
    df = df.drop(0, axis=0)
    return df

def get_new_information(df, data_columns, flg_column):
    df = df.dropna(subset=flg_column)
    df = df[data_columns]

    return df

def make_mail_status_csv(year_month_date):
    master_df = spreadsheet_to_df(os.environ.get("MASTER_SPREADSHEET_KEY"), os.environ.get("MASTER_SHEET_NAME"))
    master_df = master_df[str_to_list(os.environ.get("MASTER_USE_COLUMNS"))]
    master_df.columns = ["account_name", "mail_address"]

    # 別で読み取るデータ
    check_spreadsheet_dic = json.loads(os.environ.get("CHECK_SPREADSHEET_DIC"))
    for spreadsheet_key, use_info_dic in check_spreadsheet_dic.items():
        for use_sheet, use_columns in use_info_dic.items():
            check_df = spreadsheet_to_df(spreadsheet_key, use_sheet)
            check_df = get_new_information(check_df, use_columns[:-1], use_columns[-1])
            check_df.columns = ["account_name", "mail_address"]

            master_df = pd.concat([master_df, check_df])
    
    tmp_df = master_df[master_df["mail_address"].str.contains("、")]
    master_df = master_df[~master_df["mail_address"].str.contains("、")]
    for idx, row in tmp_df.iterrows():
        mail_address_list = row["mail_address"].split("、")
        for mail_address in mail_address_list:
            new_row = pd.Series([row["account_name"], mail_address], index=["account_name", "mail_address"])
            master_df = pd.concat([master_df, pd.DataFrame([new_row])])

    master_df = master_df.reset_index(drop=True)
    for idx, row in master_df.iterrows():
        account_name = row["account_name"]
        if account_name[-1] == "店":
            master_df.iloc[idx, 0] = master_df.iloc[idx, 0][:-1]

    master_df = master_df.drop_duplicates()

    master_df["is_send"] = "未"

    master_df.to_csv(f".//TRAN_DATA//{year_month_date}//mail_status.csv", index=False)

if __name__ == "__main__":
    # year-month-dataを作成
    year_month_date = datetime.datetime.now().strftime("%Y%m%d")
    
    make_dir(year_month_date)
    make_mail_status_csv(year_month_date)
