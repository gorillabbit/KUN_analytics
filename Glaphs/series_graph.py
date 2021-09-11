from glob import glob
import pandas as pd
import matplotlib.pyplot as plt
import os
from sklearn import linear_model
import numpy as np

src = glob('H:/Projects/basedata_KUN/*.xlsx')[-1]  # エクセルファイルからデータ取得
df_base = pd.read_excel(src, sheet_name='base')
columns = df_base.columns
group = df_base.groupby(columns[11]).groups
print(columns)

font = 'MS Gothic'
plt.style.use('ggplot')

os.makedirs('H:/Projects/Glaphs/series', exist_ok=True)
fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
clf = linear_model.LinearRegression()
for series in group:
    print(series)
    view_trans = []
    for col_num in group[series]:
        view_trans.append(df_base.iloc[col_num, 9])
    view_trans = pd.DataFrame(view_trans)
    data = view_trans.iloc[:, 0]
    nums = list(range(len(view_trans)))
    plt.bar(nums, data, color='gray')
    plt.plot(view_trans.rolling(15).mean().round(1), linewidth=3, color='#cc0000')
    plt.plot(view_trans.rolling(5).mean().round(1), linewidth=2, color='#e06666', linestyle=':')

    x = np.array([nums]).T
    y = data.values
    clf.fit(x, y)
    a = clf.coef_
    b = clf.intercept_
    plt.plot(nums, nums*a+b, color='#3c78d8', linewidth=2)
    plt.annotate('線形近似(青直線)', (0, 0), xytext=(-10, 420), textcoords="offset points", font=font, fontsize=14, color='#3c78d8')
    plt.annotate('回帰係数:'+str(a), (0, 0), xytext=(0, 405), textcoords="offset points", font=font, fontsize=12, color='#3c78d8')
    plt.annotate('切片:'+str(b), (0, 0), xytext=(0, 390), textcoords="offset points", font=font, fontsize=12, color='#3c78d8')
    plt.annotate('15本移動平均(赤線)', (0, 0), xytext=(-10, 370), textcoords="offset points", font=font, fontsize=14, color='#cc0000')
    plt.annotate('5本移動平均(薄赤点線)', (0, 0), xytext=(-10, 350), textcoords="offset points", font=font, fontsize=14, color='#e06666')

    plt.title(series, fontname=font)
    fig.savefig('H:/Projects/Glaphs/series/'+series.replace(':', ' ')+'.png')
    plt.clf()