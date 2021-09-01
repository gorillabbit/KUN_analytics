import datetime
import glob
import logging
import os

import gspread as gs
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials as SACs

nowtime = str(datetime.datetime.now()).replace(':', '-').replace('.', '-')
logging.basicConfig(filename='H:/log/make_vct' + nowtime + '.log', level=logging.DEBUG)

api_scope = ['https://www.googleapis.com/auth/spreadsheets',  # 利用する API を指定する
             'https://www.googleapis.com/auth/drive']
credentials_path = 'H:/Youtube/analytics-315310-0789825df3f6.json'  # 先ほどダウンロードした json パスを指定する
os.path.join(os.path.expanduser('~'), 'path', 'to', 'analytics-315310-0789825df3f6.json')
credentials = SACs.from_json_keyfile_name(credentials_path, api_scope)  # json から Credentials 情報を取得
gspread_client = gs.authorize(credentials)  # 認可されたクライアントを得る

ss = gspread_client.open_by_key('1-8QnVNtgva-D10P6uBgbosStPUiwq82tzcdEiaiKx8U')
s_vct = ss.get_worksheet(2)
vct = s_vct.get_all_values()
hour_df = pd.DataFrame(vct)
logging.info(hour_df)

folder_path = 'H:/Projects/basedata_KUN/vct/'
os.makedirs(folder_path, exist_ok=True)  # vctフォルダー作成
nowtime = str(datetime.datetime.now()).replace(':', '-').replace('.', '-')
filepath = folder_path + 'vct_base' + nowtime + '_raw.xlsx'

latest_vct = glob.glob(folder_path + '*_raw.xlsx')[-1]
df_latest_vct = pd.read_excel(latest_vct, index_col=0)

df_merged = pd.merge(df_latest_vct, hour_df, on=0, how='outer')
df_merged.columns = range(df_merged.shape[1])
df_merged.to_excel(filepath, sheet_name='1時間ごとの再生数')

s_vct.resize(cols=1)  # VideoID以外削除
print(s_vct)
