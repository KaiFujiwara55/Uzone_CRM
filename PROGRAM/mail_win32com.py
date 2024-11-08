# win32comを利用して、Outlook操作をプログラムによって自動化して送信する手法

import os
from dotenv import load_dotenv
load_dotenv("MASTER_DATA/.env")
import sys
import datetime
import pandas as pd
import traceback
import win32com.client
from pathlib import Path
import time
from tqdm import tqdm

def mailSend(mail_address, subject, body_text, img_path, flg):
    # OutlookAPP のインスタンス化
    outlook = win32com.client.Dispatch("Outlook.Application")
    # メールオブジェクトの作成
    mail = outlook.CreateItem(0)  # 0: メールアイテム
    mail.bodyFormat = 2 # 2: htmlメールのフォーマットを指定

    mail.to = mail_address
    mail.cc = os.environ.get("CC_MAILADDRESS")
    mail.subject = subject
    mail.HTMLBody = body_text

    # 画像を添付、添付ファイルをメールに埋め込むためにidをつける
    attachment = mail.Attachments.Add(img_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    # 送信
    if flg:
        print("メールのプレビューを確認して、問題がなければ閉じてください")
        mail.display(flg)
    else:
        try:
            mail.send
            return True
        except:
            return False

def main(year_month_date):
    try:
        # 本文をhtmlファイルをテキストで取得
        with open(os.environ.get("MAIL_BODY"), mode="rb") as g:
            html = g.read().decode("utf-8")

        # 件名を取得
        subject = open(os.environ.get("MAIL_SUBJECT").replace("{year-month-date}", year_month_date), "r", encoding="utf-8").read()
        # 部品名を取得
        parts_name = open(os.environ.get("PARTS_NAME").replace("{year-month-date}", year_month_date), "r", encoding="utf-8").read()
        # img_pathを取得
        if os.path.exists(os.environ.get("IMG").replace("{year-month-date}", year_month_date)):
            img_path = os.environ.get("IMG").replace("{year-month-date}", year_month_date)
            img_path = os.getcwd()+"\\"+img_path
        else:
            raise Exception("画像が存在しません。画像を配置してください。")

        # 送信先dfを取得
        mail_status_df = pd.read_csv(os.environ.get("MAIL_STATUS_CSV").replace("{year-month-date}", year_month_date))

        # 送信済みのものを再度送る確認
        if mail_status_df["is_send"].isin(["済"]).any():
            while True:
                x = input("送信済みのものを再送しますか？(y or n)")
                if x == "y":
                    reSendFlg = True
                    break
                elif x == "n":
                    reSendFlg = False
                    break
                else:
                    continue
        else:
            reSendFlg = False

        # 1通目はプレビュー用
        # 店名、部品名部分を置換する
        body_text = html
        body_text = body_text.replace("{店名}", "TEST店")
        body_text = body_text.replace("{部品名}", parts_name)
        mailSend("test@test.com", subject, body_text, img_path, flg=True)
        while True:
            x = input("メールの送信を開始しますか？(y or n)")
            if x == "y":
                break
            elif x == "n":
                sys.exit()
            else:
                print("「y」「n」で入力してください")

        for idx, row in tqdm(mail_status_df.iterrows(), total=len(mail_status_df)):
            account_name = row["account_name"]
            mail_address = row["mail_address"]
            status = row["is_send"]

            # 店名、部品名部分を置換する
            body_text = html
            body_text = body_text.replace("{店名}", account_name)
            body_text = body_text.replace("{部品名}", parts_name)

            # メール送信
            if type(mail_address) is str:
                if (status == "未" or reSendFlg):
                    is_send = mailSend(mail_address, subject, body_text, img_path, flg=False)
                    
                    if is_send:
                        mail_status_df.loc[idx, "is_send"] = "済"
                    else:
                        mail_status_df.loc[idx, "is_send"] = "失敗"
                    mail_status_df.to_csv(os.environ.get("MAIL_STATUS_CSV").replace("{year-month-date}", year_month_date), index=False)
                    time.sleep(5)
        
    except Exception as e:
        etype, value, tb = sys.exc_info()
        errorMessage = "".join(traceback.format_exception(etype, value, tb))
        print(errorMessage)

        # エラーが出たことを知らせ、再度実行を促す
        print("エラーのため、正常に終了することが出来ませんでした。")
        x = input("Enterを押して終了")
        sys.exit()
    
if __name__=="__main__":
    # year-month-dataを作成
    year_month_date = datetime.datetime.now().strftime("%Y%m%d")

    main(year_month_date)
