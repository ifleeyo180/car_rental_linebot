from __future__ import unicode_literals
import os
from datetime import datetime
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import pygsheets
import configparser

# Local test settings
'''
# LINE Bot 基本資料
config = configparser.ConfigParser()
config.read('config.ini')

line_bot_api = LineBotApi(config.get('line-bot', 'channel_access_token'))
handler = WebhookHandler(config.get('line-bot', 'channel_secret'))

# Google Sheets 基本資料
gc = pygsheets.authorize(service_file=config.get(
    'google-sheet', 'credentials_file'))
spreadsheet_key = config.get('google-sheet', 'spreadsheet_key')
worksheet_name = config.get('google-sheet', 'worksheet_name')
'''

google_sheet_id = os.getenv('GOOGLE_SHEET_ID')
google_sheet_creds = os.getenv('GOOGLE_SHEET_CREDENTIALS')
line_bot_channel_access_token = os.getenv('LINE_BOT_CHANNEL_ACCESS_TOKEN')
line_channel_secret = os.getenv('LINE_CHANNEL_SECRET')

worksheet_headers = ['車牌', '借用人姓名', '借用日期', '還車人姓名', '還車日期', '借用狀態']
car_database = ['ABC-123', 'XYZ-456']

# Flask 應用程式
app = Flask(__name__)

@app.route('/')
def home():
    return 'Hello, World!'

# LINE Bot Webhook 接口
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

# 借用公務車
    # LINE Bot 訊息回覆
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_input = event.message.text

    # 借用公務車
    if user_input.startswith("借車"):
        input_list = user_input.split()
        if len(input_list) != 3:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="輸入格式不正確，請輸入：借車 姓名 車牌")
            )
        else:
            car_number = input_list[2]
            user_name = input_list[1]

            if car_number not in car_database:
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text="該車輛不在可借用的車牌列表中")
                )
                return

            worksheet = gc.open_by_key(spreadsheet_key).worksheet_by_title(worksheet_name)
            data = worksheet.get_all_values()

            borrowed = False
            for row in data[1:]:
                if row[0] == car_number and row[5] == '借用中':
                    borrowed = True
                    break

            if borrowed:
                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text="該車輛已被借用")
                )
            else:
                for row in data[1:]:
                    if row[0] == '' and row[1] == '':
                        index = data.index(row) + 1

                        # 更新工作表
                        borrow_date_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        worksheet.update_value('A' + str(index), car_number)
                        worksheet.update_value('B' + str(index), user_name)
                        worksheet.update_value('C' + str(index), borrow_date_str)
                        worksheet.update_value('F' + str(index), '借用中')

                        line_bot_api.reply_message(
                            event.reply_token,
                            TextSendMessage(text="借車成功")
                        )
                        return

                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text="沒有可用的車輛")
                )

# 歸還公務車
    elif user_input.startswith("還車"):
        input_list = user_input.split()
        if len(input_list) != 3:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="輸入格式不正確，請輸入：還車 還車人姓名 車牌")
            )
            return

        return_name = input_list[1]
        car_number = input_list[2]

        worksheet = gc.open_by_key(
            spreadsheet_key).worksheet_by_title(worksheet_name)
        data = worksheet.get_all_values()

        for row in data[1:]:
            if row[0] == car_number and row[5] == '借用中':
                index = data.index(row) + 1
                return_date_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                worksheet.update_value('D' + str(index), return_name)
                worksheet.update_value('E' + str(index), return_date_str)
                worksheet.update_value('F' + str(index), '')

                line_bot_api.reply_message(
                    event.reply_token,
                    TextSendMessage(text="還車成功")
                )
                return

        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="還車失敗，請確認車牌及是否已借用")
        )


    # 查詢公務車狀態
    elif user_input.startswith("狀態"):
        car_number = user_input.split()[1]

        worksheet = gc.open_by_key(
            spreadsheet_key).worksheet_by_title(worksheet_name)
        data = worksheet.get_all_values()

        borrowed = False
        for row in data[1:]:
            if row[0] == car_number and row[5] == '借用中':
                borrowed = True
                latest_record = row
                break

        if borrowed:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text=f"{car_number} 目前狀態：{latest_record[5]}（借用人：{latest_record[1]}，借出日期：{latest_record[2]}）")
            )
        else:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text=f"{car_number} 目前沒有被借用")
            )



if __name__ == "__main__":
    app.run(debug=True, port=8000)
