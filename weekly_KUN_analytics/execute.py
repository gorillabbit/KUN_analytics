import datetime

import logging

nowtime = str(datetime.datetime.now()).replace(':', '-').replace('.', '-')
logging.basicConfig(filename='H:/log/execute'+nowtime+'.log', level=logging.DEBUG)