import glob
import shutil
from pathlib import Path

import ffmpeg
import os

import numpy as np
import pandas as pd
import textwrap

p = Path('H:/Youtube')  # フォルダを指定して色々取得しとく、動画の保存先も作成
folders = [str(x) for x in p.iterdir() if x.is_dir()]
srcs = glob.glob(folders[-2] + '/*.xlsx')  # エクセルファイルからデータ取得

background_path = 'H:/Youtube/background.mp4'
mplus_bold = 'H:/Youtube/mplus-2m-bold.ttf'
mplus_black = 'H:/Youtube/mplus-2c-black.ttf'
setting_base_base = {
    'shadowx': 1,
    'shadowy': 1,
    'shadowcolor': 'gray'
}
setting_base = {**setting_base_base, 'fontfile': mplus_bold}
setting_title = {**setting_base_base, 'fontfile': mplus_black}
box2_x = 1485
box2_h = box2_x + 20
box2_t = box2_x + 40
offset_b2 = 65
space_b2 = 10
fs_b2_h = 30
fs_b2_t = 50
fontcolor = '0x222222'
setting_b2_h = {**setting_base, "fontsize": fs_b2_h, "fontcolor": fontcolor}
setting_b2_t = {**setting_base, "fontsize": fs_b2_t, "fontcolor": fontcolor}
videos_path = folders[-2] + '/videos/koma'
txt_path = videos_path + '/outputs.txt'
v_ids = []

df_trans_this_week = pd.read_excel(srcs[0])
df_trans_last_week = pd.read_excel(srcs[1])
df_video_this_week = pd.read_excel(srcs[2])
df_video_last_week = pd.read_excel(srcs[3])
df_weekly = pd.read_excel(srcs[4])


def make_videos_video(df, option):
    os.makedirs(videos_path, exist_ok=True)
    thumbnails = glob.glob(folders[-2] + '/thumbnail' + option + '/*.png')
    vcts = glob.glob(folders[-2] + '/vct_video' + option + '/*.png')

    for i, (t, vct) in enumerate(zip(thumbnails, vcts)):
        print(t, vct)
        v_id = os.path.splitext(os.path.basename(t))[0]
        v_ids.append(v_id)
        thumbnail = ffmpeg.input(t)
        vct = ffmpeg.input(vct)
        title = textwrap.wrap(df['タイトル'][i], 45)  # 長いタイトルを分割
        title.append('')  # タイトルが短い場合にtitle[1]が無くなるの回避
        print(title)
        (
            ffmpeg
                .input(background_path, ss=0, t=5)
                .overlay(thumbnail, x=50, y=50)
                .drawbox(x=40, y=40, width=1300, height=740, color='black', t=5)  # サムネの枠
                .drawbox(x=40, y=805, width=1830, height=225, color='black', t=5)  # 下の枠
                .drawbox(x=box2_x, y=40, width=1920 - box2_x - 50, height=740, color='black', t=5)  # 上の枠
                .drawtext(x=55, y=845, text=title[0], fontsize=47, fontcolor=fontcolor, **setting_base)
                .drawtext(x=1055, y=900, text=title[1], fontsize=47, fontcolor=fontcolor, **setting_base)
                .drawtext(x=75, y=900, text='投稿日時 ' + df['投稿日時'][i], fontsize=30, fontcolor=fontcolor, **setting_base)
                .drawtext(x=75, y=940, text='動画ID　' + df['videoID'][i], fontsize=30, fontcolor=fontcolor,
                          **setting_base)
                .drawtext(x=75, y=980, text='シリーズ　' + str(df.iloc[i, 11]), fontsize=30, fontcolor=fontcolor,
                          **setting_base)
                .drawtext(x=box2_h, y=offset_b2, text='再生数', **setting_b2_h)
                .drawtext(x=box2_t, y=offset_b2 + space_b2 + fs_b2_h, text=df['再生数'][i], **setting_b2_t)
                .drawtext(x=box2_h, y=offset_b2 + space_b2 * 2 + fs_b2_h + fs_b2_t, text='高評価数', **setting_b2_h)
                .drawtext(x=box2_t, y=offset_b2 + space_b2 * 3 + fs_b2_h * 2 + fs_b2_t, text=df['高評価数'][i],
                          **setting_b2_t)
                .drawtext(x=box2_h, y=offset_b2 + space_b2 * 4 + fs_b2_h * 2 + fs_b2_t * 2, text='低評価数', **setting_b2_h)
                .drawtext(x=box2_t, y=offset_b2 + space_b2 * 5 + fs_b2_h * 3 + fs_b2_t * 2, text=df['低評価数'][i],
                          **setting_b2_t)
                .drawtext(x=box2_h, y=offset_b2 + space_b2 * 6 + fs_b2_h * 3 + fs_b2_t * 3, text='コメント数',
                          **setting_b2_h)
                .drawtext(x=box2_t, y=offset_b2 + space_b2 * 7 + fs_b2_h * 4 + fs_b2_t * 3, text=df['コメント数'][i],
                          **setting_b2_t)
                .drawtext(x=box2_h, y=offset_b2 + space_b2 * 8 + fs_b2_h * 4 + fs_b2_t * 4, text='動画の長さ',
                          **setting_b2_h)
                .drawtext(x=box2_t, y=offset_b2 + space_b2 * 9 + fs_b2_h * 5 + fs_b2_t * 4, text=df['長さ'][i],
                          **setting_b2_t)
                .drawtext(x=box2_h, y=offset_b2 + space_b2 * 10 + fs_b2_h * 5 + fs_b2_t * 5,
                          text='動画の勢い(順位)\n (半日→2日→4日)', **setting_b2_h)
                .drawtext(x=box2_t, y=offset_b2 + space_b2 * 11 + fs_b2_h * 7 + fs_b2_t * 5, text=df['伸び12時間'][i],
                          fontsize=fs_b2_t, fontcolor=fontcolor, **setting_base)
                .drawtext(x=box2_t + 88, y=offset_b2 + space_b2 * 11 + fs_b2_h * 7 + fs_b2_t * 5, text=df['伸び48時間'][i],
                          fontsize=fs_b2_t, fontcolor=fontcolor, **setting_base)
                .drawtext(x=box2_t + 176, y=offset_b2 + space_b2 * 11 + fs_b2_h * 7 + fs_b2_t * 5, text=df['伸び96時間'][i],
                          fontsize=fs_b2_t, fontcolor=fontcolor, **setting_base)
                .output(videos_path + '/' + v_id + '01.mp4')
                .run()
        )
        (
            ffmpeg
                .input(videos_path + '/' + v_id + '01.mp4', ss=0, t=5)
                .overlay(vct, x=50, y=50)
                .output(videos_path + '/' + v_id + '02.mp4').run()
        )

    vct_days = glob.glob(folders[-2] + '/vct_day' + option + '/*.png')
    for vct_date in vct_days:
        date = os.path.splitext(os.path.basename(vct_date))[0]
        vct_date = ffmpeg.input(vct_date)
        (
            ffmpeg
                .input(background_path, ss=0, t=5)
                .overlay(vct_date, x=50, y=50)
                .drawbox(x=40, y=40, width=1840, height=1000, color='black', t=5)  # サムネの枠
                .output(videos_path + '/' + date + '.mp4')
                .run()
        )


def concat_video(df, option):
    f = open(txt_path, 'w')
    group = df.groupby('投稿日').groups
    for i in group:
        print(i)
        f.write('file ' + i + '.mp4\n')
        for j in group[i]:
            filename = v_ids.pop(0)
            f.write('file ' + filename + '01.mp4\n' + 'file ' + filename + '02.mp4\n')
            print(j)
    f.close()
    ffmpeg \
        .input(txt_path, f='concat', safe=0) \
        .output(folders[-2] + '/videos/videos_1' + option + '.mp4', c='copy') \
        .run()
    shutil.rmtree(videos_path)


#make_videos_video(df_video_this_week, '')
#concat_video(df_video_this_week, '')

#make_videos_video(df_video_last_week, '_last')
#concat_video(df_video_last_week, '_last')


def make_videos_weekly():
    mask_1 = ffmpeg.input('H:/Youtube/background.png').crop(x=0, y=855, width=1920, height=34)
    mask_2 = ffmpeg.input('H:/Youtube/background.png').crop(x=1700, y=855, width=30, height=200)

    os.makedirs(videos_path, exist_ok=True)
    f = open(txt_path, 'w')
    weekly_graphs = glob.glob(folders[-2] + '/weekly/*.png')

    df_rank = df_weekly.rank(numeric_only=True, ascending=False, method='dense')
    df_rank_2 = df_video_this_week.rank(numeric_only=True, ascending=False, method='dense')
    for i, w_g in enumerate(weekly_graphs):
        graph_name = os.path.splitext(os.path.basename(w_g))[0]
        weekly_graph = ffmpeg.input(w_g)

        w_g_col = graph_name[4:]
        rank_all = []
        if w_g_col in df_rank.columns:
            col = df_rank[w_g_col]
            rank_all = pd.concat([df_weekly[col == 1][['期間', w_g_col]],
                                  df_weekly[col == 2][['期間', w_g_col]],
                                  df_weekly[col == 3][['期間', w_g_col]]])
            rank_all = pd.DataFrame(rank_all)
            rank_all.index = [1, 2, 3]

        rank_week = pd.DataFrame(index=[1, 2, 3, 4, 5], columns=['タイトル', '2'])
        rank_week['タイトル'] = 'null'
        rank_week['2'] = 'null'
        if w_g_col in df_rank_2.columns:
            col = df_rank_2[w_g_col]
            rank_week = pd.concat([df_video_this_week[col == 1][['タイトル', w_g_col]],
                                   df_video_this_week[col == 2][['タイトル', w_g_col]],
                                   df_video_this_week[col == 3][['タイトル', w_g_col]],
                                   df_video_this_week[col == 4][['タイトル', w_g_col]],
                                   df_video_this_week[col == 5][['タイトル', w_g_col]]])
            rank_week.index = [1, 2, 3, 4, 5]
        rank_week_title = pd.DataFrame(rank_week['タイトル'])
        rank_week_obj = pd.DataFrame(rank_week.iloc[:, 1])

        weekly_data = pd.DataFrame(df_weekly.iloc[-20:, [0, i + 5]])
        (
            ffmpeg
                .input(background_path, ss=0, t=5)
                .overlay(weekly_graph, x=50, y=50)
                .drawbox(x=40, y=40, width=1300, height=740, color='black', t=5)  # サムネの枠
                .drawbox(x=1370, y=40, width=1920 - 1370 - 50, height=740, color='black', t=5)  # 右の枠
                .drawtext(x=1400, y=50, text=weekly_data, fontsize=25, fontcolor=fontcolor, **setting_base)
                .drawtext(x=255, y=805, text='週のランキング', fontsize=47, fontcolor=fontcolor, **setting_base)
                .drawtext(x=107, y=855, text=rank_all, fontsize=32, fontcolor=fontcolor, **setting_base)
                .drawtext(x=1055, y=805, text='今週の動画のランキング', fontsize=47, fontcolor=fontcolor, **setting_base)
                .drawtext(x=807, y=868, text=rank_week_title, fontsize=20, fontcolor=fontcolor, **setting_base)
                .drawtext(x=1707, y=868, text=rank_week_obj, fontsize=20, fontcolor=fontcolor, **setting_base)
                .overlay(mask_1, x=0, y=855)
                .overlay(mask_2, x=1700, y=855)
                .drawtext(x=85, y=855, text='順位', fontsize=32, fontcolor=fontcolor, **setting_base)
                .drawtext(x=257, y=855, text='期間', fontsize=32, fontcolor=fontcolor, **setting_base)
                .drawtext(x=450, y=855, text=w_g_col, fontsize=32, fontcolor=fontcolor, **setting_base)
                .drawtext(x=800, y=865, text='順位', fontsize=20, fontcolor=fontcolor, **setting_base)
                .drawtext(x=1205, y=865, text='タイトル', fontsize=20, fontcolor=fontcolor, **setting_base)
                .drawtext(x=1700, y=865, text=w_g_col, fontsize=20, fontcolor=fontcolor, **setting_base)
                .output(videos_path + '/' + graph_name + '.mp4')
                .run()
        )
        f.write('file ' + graph_name + '.mp4\n')
    f.close()
    ffmpeg.input(txt_path, f='concat', safe=0).output(folders[-2] + '/videos/videos_weekly.mp4', c='copy').run()
    shutil.rmtree(videos_path)


#make_videos_weekly()

os.makedirs(videos_path, exist_ok=True)
videos = pd.DataFrame(df_video_this_week['タイトル'])
col_list = [0, 1, 2, 4, 5, 6, 7, 8, 9, 10, 11]
(
    ffmpeg
        .input(background_path, ss=0, t=5)
        .drawtext(x=50, y=50, text='動画一覧', fontsize=40, fontcolor='black', **setting_title)
        .drawtext(x=50, y=100, text=videos, fontsize=20, fontcolor='black', **setting_base)
        .drawtext(x=1100, y=50, text='指標一覧', fontsize=40, fontcolor='black', **setting_title)
        .drawtext(x=1100, y=100, text=pd.DataFrame(df_weekly.iloc[-2, col_list]), fontsize=35, fontcolor='black', **setting_base)
        .output(videos_path + '/一覧.mp4')
        .run()
)
