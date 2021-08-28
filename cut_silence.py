from pathlib import Path

import ffmpeg
import glob
import os
import soundfile as sf
import numpy as np  # 多次元配列をいっぺんに高速に計算できる
import time
import shutil

p = Path('H:/Youtube')  # フォルダを指定して色々取得しとく
folder = [str(x) for x in p.iterdir() if x.is_dir()][-2]
videos = glob.glob(folder + "/*.mp4")

for path in videos:
    basename = os.path.splitext(os.path.basename(path))[0]
    ffmpeg.input(path).output(folder + '/' + basename + ".wav", ac=1, ar=8000).run()

src_files_a = glob.glob(folder + "/*.wav")
for num, src_file in enumerate(src_files_a):

    data, samplerate = sf.read(src_file, dtype='float32')  # 音データと周波数を読み込む, dtype='float32',

    thres = 0.01  # しきい値より大きい部分を抜き出す
    amp = np.abs(data)
    b = amp > thres

    min_silence_duration = 0.2

    silences = []
    prev = 0
    entered = 0
    for i, v in enumerate(b):
        if prev == 1 and v == 0:  # enter silence
            entered = i
        if prev == 0 and v == 1:  # exit silence
            duration = (i - entered) / samplerate
            if duration > min_silence_duration:
                silences.append({"from": entered, "to": i})
                entered = 0
        prev = v
    if 0 < entered < len(b):
        silences.append({"from": entered, "to": len(b)})

    min_keep_duration = 0.2
    cut_blocks = []
    while 1:  # 無音期間の間が0.2秒以下の場合、そこもカットする
        if len(silences) == 1:
            cut_blocks.append(silences[0])
            break
        pre_s = silences[0]  # 最初の無音期間はスキップして、0番目と1番目から始める
        for i, s in enumerate(silences):
            if i == 0:
                continue
            interval = (s["from"] - pre_s["to"]) / samplerate  # 無音の間の長さ
            if interval < min_keep_duration:
                pre_s["to"] = s["to"]  # 2つの期間を連結
            else:
                cut_blocks.append(pre_s)  # 連結したのをカットする配列に突っ込む
                pre_s = s  # 次のpre_sにsを代入する
        cut_blocks.append(pre_s)
        break

    keep_blocks = []
    for i, block in enumerate(cut_blocks):
        if i == 0 and block["from"] > 0:
            keep_blocks.append({"from": 0, "to": block["from"]})
        if i > 0:
            prev = cut_blocks[i - 1]
            keep_blocks.append({"from": prev["to"], "to": block["from"]})
        if i == len(cut_blocks) - 1 and block["to"] < len(data):
            keep_blocks.append({"from": block["to"], "to": len(data)})
    filename = os.path.basename(videos[num])
    out_dir = os.path.join(folder + "/" + str((time.time())))  # ffmpegの仕様でスラッシュかバックスラッシュか円記号かでエラーでるから注意
    os.mkdir(out_dir)

    txt_path = out_dir + '/outputs.txt'
    f = open(txt_path, 'w')
    xxx = []
    for i, block in enumerate(keep_blocks):
        fr = block["from"] / samplerate
        to = block["to"] / samplerate
        duration = to - fr
        xxx.append([fr, to])
        out_path = os.path.join(out_dir + '/' + str(i) + '.mp4')
        ffmpeg.input(videos[num], ss=fr, t=duration).output(out_path, vcodec="copy", acodec="copy").run()
        f.write('file ' + str(i) + '.mp4\n')
    f.close()
    ffmpeg.input(txt_path, f='concat', safe=0).output(folder + "/" + 'short_' + filename, vcodec="copy", acodec="copy").run()
    print(filename)
    print(keep_blocks)
    #shutil.rmtree(out_dir)
    os.remove(src_file)
    print(xxx)
    exit()
