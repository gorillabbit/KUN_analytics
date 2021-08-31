import datetime
import glob
import os
import gspread as gs
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials as SACs
import logging

nowtime = str(datetime.datetime.now()).replace(':', '-').replace('.', '-')
logging.basicConfig(filename='H:/log/make_daily_view' + nowtime + '.log', level=logging.DEBUG)

api_scope = ['https://www.googleapis.com/auth/spreadsheets',  # 利用する API を指定する
             'https://www.googleapis.com/auth/drive']
credentials_path = 'H:/Youtube/analytics-315310-0789825df3f6.json'  # 先ほどダウンロードした json パスを指定する
os.path.join(os.path.expanduser('~'), 'path', 'to', 'analytics-315310-0789825df3f6.json')
credentials = SACs.from_json_keyfile_name(credentials_path, api_scope)  # json から Credentials 情報を取得
gspread_client = gs.authorize(credentials)  # 認可されたクライアントを得る

ss = gspread_client.open_by_key('1-8QnVNtgva-D10P6uBgbosStPUiwq82tzcdEiaiKx8U')
s_daily = ss.get_worksheet(1)
daily_df = pd.DataFrame(s_daily.get_all_values())
print(daily_df)
logging.info(daily_df)

folder_path = 'H:/Youtube/basedata/daily/'
os.makedirs(folder_path, exist_ok=True)  # dailyフォルダー作成
nowtime = str(datetime.datetime.now()).replace(':', '-').replace('.', '-')
filepath = folder_path + 'daily_base' + nowtime + '.xlsx'

latest_daily = glob.glob(folder_path + '*.xlsx')[-1]
df_latest_daily = pd.read_excel(latest_daily, index_col=0)
print(df_latest_daily)

df_merged = pd.merge(df_latest_daily, daily_df, on=0, how='outer')
df_merged.columns = range(df_merged.shape[1])
df_merged.to_excel(filepath, sheet_name='1時間ごとの再生数', index=False)

# s_daily.resize(cols=1)  # VideoID以外削除
