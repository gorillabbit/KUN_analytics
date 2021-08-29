import datetime

import download_sp
import make_weekly_xlsx
import download_thumbnail
import Graph
import makepptx
import logging

nowtime = str(datetime.datetime.now()).replace(':', '-').replace('.', '-')
logging.basicConfig(filename='H:/log/execute'+nowtime+'.log', level=logging.DEBUG)