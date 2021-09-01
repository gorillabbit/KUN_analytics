import glob
import sklearn
import pandas as pd

latest_vct = glob.glob('H:/Projects/basedata_KUN/vct/*.xlsx')[-1]
print(latest_vct)
df_latest_vct = pd.read_excel(latest_vct, index_col=0, sheet_name='Sheet')

print(df_latest_vct)