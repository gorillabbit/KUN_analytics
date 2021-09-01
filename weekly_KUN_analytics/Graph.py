import os
from glob import glob
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

p = Path('H:/Projects/週刊KUN分析')  # フォルダを指定して色々取得しとく
folder = [str(x) for x in p.iterdir() if x.is_dir()][-1]
srcs = glob(folder + '/*.xlsx')  # エクセルファイルからデータ取得

df_trans_this = pd.read_excel(srcs[0])
df_trans_last = pd.read_excel(srcs[1])
df_videos_this = pd.read_excel(srcs[2])
df_videos_last = pd.read_excel(srcs[3])
df_weekly = pd.read_excel(srcs[4])

font = 'MS Gothic'
x_axis = df_trans_this.columns.values[1:]
x_axis = list(map(lambda x: int(x) - 1, x_axis))[::10]

plt.style.use('ggplot')
fig = plt.figure(figsize=(18.2, 9.8), tight_layout=True)
ax = fig.add_subplot(xticks=x_axis)


def make_daily_graph(df_v, df_vct, option):
    group = df_v.groupby('投稿日').groups
    print(group)

    vtc_path = folder + '/vct_day' + option
    os.makedirs(vtc_path, exist_ok=True)

    for date in group:
        print(date)
        for column_num in group[date]:
            row = df_vct.iloc[column_num, 1:]
            v_id = df_v.at[column_num, 'タイトル']
            ax.plot(row, label=v_id, linewidth=3)

        plt.legend(bbox_to_anchor=(1, -0.07), loc='upper right', prop={'family': font})
        plt.title('再生数推移-' + date, fontname=font)
        plt.xlabel('投稿からの時間(時間)', fontname=font)
        plt.ylabel('再生数(回)', )
        date_file = date.replace('/', '-')
        fig.savefig(vtc_path + '/' + date_file + '.png')
        plt.cla()


make_daily_graph(df_videos_this, df_trans_this, '')
make_daily_graph(df_videos_last, df_trans_last, '_last')

# 　こっから動画ごとのグラフ
basedata_folder_path = 'H:/Projects/basedata_KUN/'
basedata = glob(basedata_folder_path + '*.xlsx')[-1]

df_n = pd.read_excel(basedata, sheet_name='推移', index_col=0)[-200:]  # 直近200のデータ
quantile25 = df_n.quantile(q=[0.25], numeric_only=True)
quantile75 = df_n.quantile(q=[0.75], numeric_only=True)


def graph_video(df_v, df_vc_t, option):

    fig = plt.figure(figsize=(12.8, 7.2), tight_layout=True)
    ax = fig.add_subplot(xticks=x_axis)
    for j, v_id_v in enumerate(df_v['videoID']):  # グラフを作成

        print(v_id_v)
        length = len(df_vc_t.iloc[1, 1:].dropna()) + 50
        x = list(range(0, length))  # fill_betweenのためのx

        ax.plot(df_vc_t.iloc[j, 1:], linewidth=3)  # 目的の動画のグラフをプロット
        ax.fill_between(x=x, y1=quantile25.loc[0.25, :length - 1], y2=quantile75.loc[0.75, :length - 1],
                        color='darkgray', alpha=0.5)

        plt.xlabel('投稿からの時間(時間)', fontname=font)
        plt.ylabel('再生数(回)', fontname=font)
        j = ('000' + str(j))[-3:]

        vtc_v_path = folder + '/vct_video' + option
        os.makedirs(vtc_v_path, exist_ok=True)
        fig.savefig(vtc_v_path + '/' + j + '_' + v_id_v + '.png')
        plt.cla()


graph_video(df_videos_this, df_trans_this, '')  # 今週の
graph_video(df_videos_last, df_trans_last, '_last')  # 先週の

# こっから週間の推移のグラフ

weekly_path = folder + '/weekly'
os.makedirs(weekly_path, exist_ok=True)
x_ticks = list(df_weekly['期間'][:-1].str[9:])
df_rank = df_weekly.rank(numeric_only=True, ascending=False, method='dense')


def add_rank_to_bar(bar_graph, axes, col_name, color, f_size=25, y_offset=-25, y_position=100000, mod=''):  # 順位の印をつける
    for i in range(5):
        ranks = df_rank[df_rank[col_name] == i+1].index
        for rank in ranks:  # 順位が同じやつが複数ある場合全てに記しするため
            rank_x = bar_graph[rank].get_x()
            rank_y = bar_graph[rank].get_height()
            if y_position != 100000:
                rank_y = y_position
            axes.annotate(str(i+1)+mod, xy=(rank_x, rank_y), xytext=(3, y_offset), fontsize=f_size, color=color, textcoords="offset points")


def add_change_and_so_on(bar_graph, axes, place='center', color='white'):
    for i in range(1, len(bar_graph)):  # 前からの変化量と、変化量(ポイント)の記入
        change = int(bar_graph[-i].get_height() - bar_graph[-(i+1)].get_height())
        change_rate = int(((bar_graph[-i].get_height() / bar_graph[-(i+1)].get_height())-1)*100)
        y = bar_graph[-i].get_height()/2
        if place == 'top':
            y = bar_graph[-i].get_height()
        axes.annotate('({:+}'.format(change)+')\n({:+}'.format(change_rate)+'pt)',
                      xy=(bar_graph[-i].get_x()+bar_graph[-i].get_width()/2, y),
                      xytext=(0, -40), fontsize=14, color=color, ha='center', textcoords="offset points")


fig_size_x = 23.5
fig_size_y = 9.0

fig_w = plt.figure(figsize=(fig_size_x, fig_size_y), tight_layout=True)  # 高評価と低評価
ax_rate_1 = fig_w.add_subplot()
ax_rate_2 = ax_rate_1.twinx()
ax_rate_1.plot(df_weekly.iloc[:-1, 8], color='#e06666', linewidth=3)
ax_rate_1.plot(df_weekly.iloc[:-1, 10]*20, color='#6fa8dc', linewidth=3)
ax_rate_1.bar(x_ticks, df_weekly.iloc[:-1, 14]/75, color='#a4c2f4', zorder=0, alpha=0.7)
ax_rate_2.plot(df_weekly.iloc[:-1, 20], color='#8e7cc3', linewidth=2, linestyle='--')
like_c = ax_rate_1.bar(x_ticks, df_weekly.iloc[:-1, 8], alpha=0, width=0.2)
unlike_c = ax_rate_1.bar(x_ticks, df_weekly.iloc[:-1, 10]*20, alpha=0, width=0.2)
ratio = ax_rate_2.bar(x_ticks, df_weekly.iloc[:-1, 20], alpha=0, width=0.2)
add_rank_to_bar(like_c, ax_rate_1, '高評価数', '#e06666', y_offset=21, f_size=20, y_position=0)
add_rank_to_bar(unlike_c, ax_rate_1, '低評価数', '#6fa8dc', y_offset=1, f_size=20, y_position=0)
add_rank_to_bar(ratio, ax_rate_2, '高評価-低評価比率', '#8e7cc3', y_offset=41, f_size=20, y_position=0)
ax_rate_1.set_ylim(0)
ax_rate_2.set_ylim(0)
ax_rate_1.set_ylabel('高評価数(個),低評価数x20(個)', fontname=font)
ax_rate_2.set_ylabel('高評価数/低評価数', fontname=font)
ax_rate_2.grid(False)
ax_rate_1.grid(color='black', linestyle=':')
plt.title('高評価数と低評価数(と再生数)', fontname=font)
fig_w.savefig(weekly_path + '/002_高評価数と低評価数.png', transparent=True)
plt.close('all')


def make_weekly_graph_1(name, color, col_num_list, number):
    fig_g = plt.figure(figsize=(fig_size_x, fig_size_y), tight_layout=True)
    ax1 = fig_g.add_subplot(211)
    ax2 = fig_g.add_subplot(212)
    ax1_2 = ax1.twinx()
    ax2_2 = ax2.twinx()
    ax1.plot(x_ticks, df_weekly.iloc[:-1, col_num_list[0]], color=color, label=name, linewidth=3, linestyle='-')
    ax1.plot(df_weekly.iloc[:-1, col_num_list[1]]*30, color=color, label=name+'_平均x30', linewidth=3, linestyle='--')
    ax2.plot(x_ticks, df_weekly.iloc[:-1, col_num_list[2]], color=color, label=name+'_1000再生あたり', linewidth=3, linestyle=':')
    quantity = ax1_2.bar(x_ticks, df_weekly.iloc[:-1, col_num_list[0]], alpha=0, width=0.2)
    average = ax1_2.bar(x_ticks, df_weekly.iloc[:-1, col_num_list[1]]*30, alpha=0, width=0.2)
    thousand_per = ax2_2.bar(x_ticks, df_weekly.iloc[:-1, col_num_list[2]], alpha=0, width=0.2)
    add_rank_to_bar(quantity, ax1_2, name, color, y_offset=23, f_size=20, y_position=0)
    add_rank_to_bar(average, ax1_2, name+'_平均', color, y_offset=5, f_size=17, y_position=0, mod='(Avg)')
    add_rank_to_bar(thousand_per, ax2_2, name+'(再生数比x1000)', color, y_offset=5, f_size=20, y_position=0)
    ax1_2.axis('off')
    ax2_2.axis('off')
    ax1.set_ylabel('高評価数(破線：平均x30)', fontname=font)
    ax2.set_ylabel('1000再生あたり高評価数', fontname=font)
    ax1.grid(color='black', linestyle=':')
    ax2.grid(color='black', linestyle=':')
    plt.title(name, fontname=font)
    fig_g.savefig(weekly_path+'/10'+str(number)+'_'+name+'.png', transparent=True)
    plt.close('all')


make_weekly_graph_1('高評価数', '#e06666', [8, 9, 17], 1)
make_weekly_graph_1('低評価数', '#6fa8dc', [10, 11, 18], 2)
make_weekly_graph_1('コメント数', '#93c47d', [12, 13, 19], 3)

fig_w = plt.figure(figsize=(fig_size_x, fig_size_y), tight_layout=True)
ax_w_1 = fig_w.add_subplot()  # 再生数
ax_w_2 = ax_w_1.twinx()
total_view = ax_w_1.bar(x_ticks, df_weekly.iloc[:-1, 16], alpha=0.9, color='#a4c2f4', edgecolor="#4a86e8", linewidth=2)
weekly_view = ax_w_1.bar(x_ticks, df_weekly.iloc[:-1, 14], alpha=0.9, color='#4a86e8')
ax_w_2.plot(df_weekly.iloc[:-1, 15], color='#e06666', linewidth=3)
view_c_ave = ax_w_2.bar(x_ticks, df_weekly.iloc[:-1, 15], alpha=0, width=0.2)
ax_w_1.bar_label(total_view, color='#4a86e8', fmt='%.0f', fontsize=15)
ax_w_1.bar_label(weekly_view, label_type='center', color='white', fmt='%.0f', fontsize=15)
add_rank_to_bar(total_view, ax_w_1, '総再生数', '#4a86e8')
add_rank_to_bar(weekly_view, ax_w_1, '再生数', 'white')
add_rank_to_bar(view_c_ave, ax_w_2, '再生数_平均', '#990000', y_offset=10, f_size=20, y_position=100001, mod='(Avg)')
add_change_and_so_on(total_view, ax_w_1, place='top', color='#4a86e8')
add_change_and_so_on(weekly_view, ax_w_1)
ax_w_1.set_ylabel('再生数', fontname=font)
ax_w_2.set_ylabel('平均再生数', fontname=font)
ax_w_2.set_ylim(100000, 300000)
ax_w_2.grid(False)
ax_w_1.grid(color='black', linestyle=':')
plt.title('週間総再生数とその内その週に投稿された動画の割合、平均再生数', fontname=font)
fig_w.savefig(weekly_path + '/001_再生数.png', transparent=True)
plt.close('all')

fig_w = plt.figure(figsize=(fig_size_x, fig_size_y), tight_layout=True)
ax_w_1 = fig_w.add_subplot()  # 長さ
ax_w_2 = ax_w_1.twinx()
duration = ax_w_1.bar(x_ticks, df_weekly.iloc[:-1, 6], alpha=0.9, color='#a4c2f4')
ax_w_2.plot(df_weekly.iloc[:-1, 7],  color='#e06666', linewidth=3)
avg_duration = ax_w_2.bar(x_ticks, df_weekly.iloc[:-1, 7], alpha=0, width=0.2)
duration[-1].set_color("#4a86e8")
ax_w_1.bar_label(duration, label_type='center', color='white', fmt='%.2f', fontsize=15)
add_rank_to_bar(duration, ax_w_1, "長さ(分)", 'white')
add_rank_to_bar(avg_duration, ax_w_2, '長さ(分)_平均', '#e06666', y_offset=5, f_size=20, y_position=0, mod='(Avg)')
add_change_and_so_on(duration, ax_w_1)
ax_w_1.set_ylim(bottom=0)
ax_w_2.set_ylim(bottom=0)
ax_w_2.grid(False)
ax_w_1.set_ylabel('動画長(分)', fontname=font)
ax_w_2.set_ylabel('平均動画長(分)', fontname=font)
ax_w_1.grid(color='black', linestyle=':')
plt.title('動画の長さ', fontname=font)
fig_w.savefig(weekly_path + '/200_長さ.png', transparent=True)

fig_w = plt.figure(figsize=(fig_size_x, 12.37), tight_layout=True)  # 動画数
ax_w = fig_w.add_subplot()
g = ax_w.bar(x_ticks, df_weekly.iloc[:-1, 5], alpha=0.9, color='#a4c2f4')
g[-1].set_color("#4a86e8")
ax_w.set_ylim(bottom=0)
ax_w.bar_label(g, label_type='center', color='white', fmt='%.2f', fontsize=15)
add_rank_to_bar(g, ax_w, '動画数(個)', 'white')
add_change_and_so_on(g, ax_w)
ax_w.grid(color='black', linestyle=':')
plt.title('動画数(個)', fontname=font)
fig_w.savefig(weekly_path + '/000_動画数.png', transparent=True)
plt.close('all')
