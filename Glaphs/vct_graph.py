import matplotlib.pyplot as plt
import pandas as pd

skiprows = list(range(1,1000))
df_base = pd.read_excel('H:/Projects/Glaphs/basedata_2.xlsx', sheet_name='推移', skiprows=skiprows, index_col=0)
font = 'MS Gothic'
plt.style.use('ggplot')
fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)

df_vct = df_base.iloc[-300:, :12]
for i in range(len(df_vct)):
    plt.plot(df_vct.iloc[i, 1:], alpha=0.4, linewidth=1, color='#e06666')
fig.savefig('H:/Projects/Glaphs/12時間までの再生数推移.png')
plt.close('all')

fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
df_vct = df_base.iloc[-300:, :48]
for i in range(len(df_vct)):
    plt.plot(df_vct.iloc[i, 1:], alpha=0.4, linewidth=1, color='#e06666')
fig.savefig('H:/Projects/Glaphs/48時間までの再生数推移.png')
plt.close('all')

fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
df_vct = df_base.iloc[-300:, :96]
for i in range(len(df_vct)):
    plt.plot(df_vct.iloc[i, 1:], alpha=0.4, linewidth=1, color='#e06666')
fig.savefig('H:/Projects/Glaphs/96時間までの再生数推移.png')
plt.close('all')

fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
df_vct = df_base.iloc[-300:, :144]
for i in range(len(df_vct)):
    plt.plot(df_vct.iloc[i, 1:], alpha=0.4, linewidth=1, color='#e06666')
fig.savefig('H:/Projects/Glaphs/144時間までの再生数推移.png')
plt.close('all')

fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
df_vct = df_base.iloc[-300:, :192]
for i in range(len(df_vct)):
    plt.plot(df_vct.iloc[i, 1:], alpha=0.4, linewidth=1, color='#e06666')
fig.savefig('H:/Projects/Glaphs/192時間までの再生数推移.png')
plt.close('all')

fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
df_vct = df_base.iloc[-300:, :240]
for i in range(len(df_vct)):
    plt.plot(df_vct.iloc[i, 1:], alpha=0.4, linewidth=1, color='#e06666')
fig.savefig('H:/Projects/Glaphs/240時間までの再生数推移.png')
plt.close('all')

