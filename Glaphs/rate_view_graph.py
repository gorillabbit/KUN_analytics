import matplotlib.pyplot as plt
import pandas as pd
from glob import glob

skiprows = list(range(1, 8500))
src = glob('H:/Projects/basedata_KUN/*.xlsx')[-1]
df_base =  pd.read_excel(src, sheet_name='base', skiprows=skiprows)
df_vct =  pd.read_excel(src, sheet_name='推移', index_col=0)
font = 'MS Gothic'
plt.style.use('ggplot')
fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)

df_base['高評価数(再生数比x1000)'] = df_base['高評価数'] / df_base['再生数'] * 1000
df_base['低評価数(再生数比x1000)'] = df_base['低評価数'] / df_base['再生数'] * 1000

df_sorted = df_base.sort_values('低評価数(再生数比x1000)')

for i in range(40):
    print(i)
    row = df_vct[df_vct['videoID'] == df_sorted.iloc[-i+1, 0]]
    print(row)
    if not row.empty:
        plt.plot(row.iloc[0, 1:], color='#6fa8dc')
    row = df_vct[df_vct['videoID'] == df_sorted.iloc[i, 0]]
    print(row)
    if not row.empty:
        plt.plot(row.iloc[0, 1:], color='#e06666')
plt.title('低評価数(1000再生あたり)が多い動画（青）と少ない動画（赤）', font=font)
plt.show()
fig.savefig('H:/Projects/Glaphs/低評価数(1000再生あたり)が多い動画（青）と少ない動画（赤）.png')
plt.close('all')

df_sorted = df_base.sort_values('高評価数(再生数比x1000)')
fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
for i in range(40):
    print(i)
    row = df_vct[df_vct['videoID'] == df_sorted.iloc[-i+1, 0]]
    print(row)
    if not row.empty:
        plt.plot(row.iloc[0, 1:], color='#e06666')
    row = df_vct[df_vct['videoID'] == df_sorted.iloc[i, 0]]
    print(row)
    if not row.empty:
        plt.plot(row.iloc[0, 1:], color='#6fa8dc')

plt.title('高評価数(1000再生あたり)が多い動画（赤）と少ない動画（青）', font=font)
plt.show()
fig.savefig('H:/Projects/Glaphs/高評価数(1000再生あたり)が多い動画（赤）と少ない動画（青）.png')