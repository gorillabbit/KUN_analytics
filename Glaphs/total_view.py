from glob import glob
import pandas as pd
import matplotlib.pyplot as plt

src = glob('H:/Projects/basedata_KUN/*.xlsx')[-1]  # エクセルファイルからデータ取得
df_base = pd.read_excel(src, sheet_name='base')

font = 'MS Gothic'
plt.style.use('ggplot')
fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
ax = fig.add_subplot()

ax.bar(list(range(len(df_base['再生数']))), df_base['再生数'], alpha=0.5)
ax.plot(df_base['再生数'].rolling(15).mean().round(1), linewidth=0.5, color='#cc0000', linestyle=':')
ax.plot(df_base['再生数'].rolling(100).mean().round(1), linewidth=1, color='#cc0000')
ax.set_ylim(0, 2000000)
fig.savefig('H:/Projects/Glaphs/全ての動画の再生数(時系列).png')