import os
import urllib.request
from glob import glob
from pathlib import Path

import cv2
import numpy as np
import pandas as pd
from PIL import Image

p = Path('H:/Youtube')  # フォルダを指定して色々取得しとく、動画の保存先も作成
folders = [str(x) for x in p.iterdir() if x.is_dir()]
srcs = glob(folders[-2] + '/*.xlsx')  # エクセルファイルからデータ取得


def download_thumbnail(option, df):
    tn_folder = folders[-2] + '/thumbnail' + option
    os.makedirs(tn_folder, exist_ok=True)
    print(df)
    thumbnail = df['サムネイル']
    v_id = df['videoID']
    print(thumbnail)
    for i, t in enumerate(thumbnail):
        num = ('000' + str(i))[-3:]
        dst_path = tn_folder + '/' + num + '_' + v_id[i] + '.png'
        try:
            with urllib.request.urlopen(t) as pic_thumb, open(dst_path, mode='wb') as local_file:
                local_file.write(pic_thumb.read())
        except urllib.error.HTTPError:
            blank = np.zeros((710, 1280, 3))
            cv2.imwrite(dst_path, blank)
    tns = glob(tn_folder + '/*')
    for tn in tns:
        Image.open(tn).resize((1280, 720)).save(tn)


download_thumbnail('', pd.read_excel(srcs[2]))  # 今週
download_thumbnail('_last', pd.read_excel(srcs[3]))  # 先週
