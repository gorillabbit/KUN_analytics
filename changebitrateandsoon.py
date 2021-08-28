import glob
import os
import ffmpeg

path = 'H:/Youtube/output'
os.makedirs(path, exist_ok=True)
videos = glob.glob('H:/Youtube/*.mp4')
print(videos)

for video in videos:
    print(video[11:])
    ffmpeg.input(video).output(path +'/' + video[11:], r=24.24).run()