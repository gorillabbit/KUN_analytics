from glob import glob
import matplotlib.pyplot as plt
import pandas as pd

def change_length(x):
    element = str(x).split(':')
    result = int(element[0]) * 60 + int(element[1]) + int(element[2]) / 60
    result = round(result, 2)
    return result


src = glob('H:/Projects/basedata_KUN/*.xlsx')[-1]  # エクセルファイルからデータ取得
skiprows = list(range(1,6000))
df_base = pd.read_excel(src, sheet_name='base', skiprows=skiprows)
df_base['高評価数(再生数比x1000)'] = df_base['高評価数']/df_base['再生数'] * 1000
df_base['低評価数(再生数比x1000)'] = df_base['低評価数']/df_base['再生数'] * 1000
df_base['コメント数(再生数比x1000)'] = df_base['コメント数']/df_base['再生数'] * 1000
df_base['高評価-低評価比率'] = df_base['高評価数']/df_base['低評価数']
df_base['長さ(分)'] = df_base['長さ'].apply(change_length)
df_base['投稿日'] = df_base['投稿日時'].apply(lambda x: x[0:10].replace('/', '-'))

print(df_base)
font = 'MS Gothic'
plt.style.use('ggplot')
fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
plt.title('再生数と1000再生あたりの高評価数の関係', fontname=font)
plt.scatter(df_base['再生数'], df_base['高評価数(再生数比x1000)'], s=2)
fig.savefig('H:/Projects/Glaphs/like_rate_and_view.png')
#plt.xscale("log")
plt.show()

