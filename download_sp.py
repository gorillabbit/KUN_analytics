import datetime
import glob
import os

import gspread as gs
import openpyxl
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials as SACs

api_scope = ['https://www.googleapis.com/auth/spreadsheets',  # 利用する API を指定する
             'https://www.googleapis.com/auth/drive']
credentials_path = 'H:/Projects/週刊KUN分析/analytics-315310-0789825df3f6.json'  # 先ほどダウンロードした json パスを指定する
os.path.join(os.path.expanduser('~'), 'path', 'to', 'analytics-315310-0789825df3f6.json')
credentials = SACs.from_json_keyfile_name(credentials_path, api_scope)  # json から Credentials 情報を取得
gspread_client = gs.authorize(credentials)  # 認可されたクライアントを得る
ss = gspread_client.open_by_key('1-8QnVNtgva-D10P6uBgbosStPUiwq82tzcdEiaiKx8U')

s_base = ss.get_worksheet(0)
s_daily = glob.glob('H:/Projects/basedata_KUN/daily/*.xlsx')[-1]
s_hour = glob.glob('H:/Projects/basedata_KUN/vct/*.xlsx')[-1]

nowtime = str(datetime.datetime.now()).replace(':', '-').replace('.', '-')
wb = openpyxl.Workbook()
wb_sheet = wb.active
wb_sheet.title = 'Blank'
folder_path = 'H:/Projects/basedata_KUN/'
os.makedirs(folder_path, exist_ok=True)
filepath = folder_path + nowtime + '.xlsx'
wb.save(filepath)

base = s_base.get_all_records()  # シートからデータを入手し成形
base_df = pd.DataFrame.from_records(base, columns=base[0].keys(), index='videoID')
daily_df = pd.read_excel(s_daily, sheet_name='1時間ごとの再生数', skiprows=0)
hour_df = pd.read_excel(s_hour, sheet_name='Sheet')

with pd.ExcelWriter(filepath, engine="openpyxl", mode='a') as writer:  # エクセルファイルに記入
    base_df.to_excel(writer, sheet_name='base')
    daily_df.to_excel(writer, sheet_name='日間再生数', index=False, header=False)
    hour_df.to_excel(writer, sheet_name='推移')