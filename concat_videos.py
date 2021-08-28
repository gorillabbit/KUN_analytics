import os
from pathlib import Path
import ffmpeg
import shutil

p = Path('H:/Youtube')  # フォルダを指定して色々取得しとく、動画の保存先も作成
folder = [str(x) for x in p.iterdir() if x.is_dir()][-2]

os.makedirs(folder + '/videos', exist_ok=True)

#shutil.copy(folder + '/opening_slide_0.mp4', folder + '/videos/01_opening_slide.mp4')
#shutil.copy(folder + '/weekly.mp4', folder + '/videos/04_weekly.mp4')
#shutil.copy(folder + '/daily.mp4', folder + '/videos/06_daily.mp4')
#shutil.copy(folder + '/daily_last.mp4', folder + '/videos/08_daily_last.mp4')

shutil.copy('H:/Youtube/opening.mp4', folder + '/videos/02_opening.mp4')
shutil.copy('H:/Youtube/weekly_info.mp4', folder + '/videos/03_weekly_info.mp4')
shutil.copy('H:/Youtube/v_info.mp4', folder + '/videos/05_v_info.mp4')
shutil.copy('H:/Youtube/v_info_last.mp4', folder + '/videos/07_v_info_last.mp4')

shutil.copy('H:/Youtube/Crazy_Blues.mp3', folder + '/videos/00_Crazy_Blues.mp3')

txt_path = folder + '/videos/videos.txt'
f = open(txt_path, 'w')
f.write('file opening_slide.mp4\n'
        'file opening.mp4\n'
        'file weekly_info.mp4\n'
        'file weekly.mp4\n'
        'file v_info.mp4\n'
        'file daily.mp4\n'
        'file v_info_last.mp4\n'
        'file daily_last.mp4')
f.close()
#ffmpeg.input(txt_path, f='concat', safe=0).output(folder + "/video.mp4", r=24.24).run()


