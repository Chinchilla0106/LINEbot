from flask import Flask, request, abort

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
)
import os
import pandas as pd
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
from dateutil.tz import gettz

def auth():
    SP_CREDENTIAL_FILE = "YOUR_JSON_FILE"
    SP_SCOPE = [
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"
                ]
    SP_SHEET_KEY = "YOUR_SHEET_KEY"
    SP_SHEET = "YOUR_SHEET"

    credentials = ServiceAccountCredentials.from_json_keyfile_name(SP_CREDENTIAL_FILE, SP_SCOPE)
    gc = gspread.authorize(credentials)

    worksheet = gc.open_by_key(SP_SHEET_KEY).worksheet(SP_SHEET)
    return worksheet

####出勤####
def punch_in():
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    timestamp = datetime.datetime.now().astimezone(gettz('Asia/Tokyo'))
    date = timestamp.strftime('%Y/%m/%d')
    punch_in = timestamp.strftime('%H:%M')
    df.loc[""] = ["", "", "", ""]
    df.iloc[-1, 0] = date
    df.iloc[-1, 1] = punch_in

    worksheet.update([df.columns.values.tolist()] + df.values.tolist())



####退勤####
def punch_out():
    worksheet = auth()#もう一度更新された値をもってくる
    df = pd.DataFrame(worksheet.get_all_records())#これももう一度
    timestamp = datetime.datetime.now().astimezone(gettz('Asia/Tokyo'))
    punch_out = timestamp.strftime('%H:%M')

    df.iloc[-1, 2] = punch_out
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

def work_time():
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())
    shukkin = df.iloc[-1, 1]
    shukkin_H = int(shukkin[:2])
    shukkin_M = int(shukkin[3:])
    shukkin_seconds = datetime.timedelta(hours=shukkin_H, minutes=shukkin_M)

    taikin = df.iloc[-1, 2]
    taikin_H = int(taikin[:2])
    taikin_M = int(taikin[3:])
    taikin_seconds = datetime.timedelta(hours=taikin_H, minutes=taikin_M)
    
    difference = (taikin_seconds - shukkin_seconds).seconds / 60
    minute = divmod(difference, 60)
    H = round(minute[0])
    M = round(minute[1])
    difference = (str(H) + ":" + str(M))
    df.iloc[-1, 3] = difference
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    
    
app = Flask(__name__)
#環境変数を入れたほうがいい
YOUR_CHANNEL_ACCESS_TOKEN=""
YOUR_CHANNEL_SECRET = ""

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)
@app.route("/")
def hello_world():
    return "hello world"

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'

#リプライメッセージの部分
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    if event.message.text == '出勤':
        punch_in()
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="出勤登録完了しました！"))
        
    elif event.message.text == '退勤':
        punch_out()            
        work_time()
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text='退勤登録完了しました。\nお疲れ様でした！！'))
            
    else:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="こちらは出退勤を管理するbotです。\n「出勤」か「退勤」と入力してください。"))


if __name__ == "__main__":
    port = os.getenv("PORT")
    app.run(host="0.0.0.0", port=port)