import os
from pathlib import Path
import shutil

p = Path('H:/Youtube')  # フォルダを指定して色々取得しとく、動画の保存先も作成
folder = [str(x) for x in p.iterdir() if x.is_dir()][-2]

os.makedirs(folder + '/videos', exist_ok=True)

shutil.copy('H:/Youtube/opening.mp4', folder + '/videos/02_opening.mp4')
shutil.copy('H:/Youtube/weekly_info.mp4', folder + '/videos/03_weekly_info.mp4')
shutil.copy('H:/Youtube/v_info.mp4', folder + '/videos/05_v_info.mp4')
shutil.copy('H:/Youtube/v_info_last.mp4', folder + '/videos/07_v_info_last.mp4')
shutil.copy('H:/Youtube/Crazy_Blues.mp3', folder + '/videos/00_Crazy_Blues.mp3')
