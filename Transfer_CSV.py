import time
import os
import shutil
from datetime import datetime

today = datetime.now()
SECONDS_IN_DAY = 24 * 60 * 60

src = "C:\Python\SQL"
dst = "C:\Python\SQL\Line_Data_" + today.strftime('%d_%m_%Y')

if not os.path.exists(dst):
    os.makedirs(dst)

now = time.time()
before = now - SECONDS_IN_DAY

def last_mod_time(fname):
    return os.path.getmtime(fname)

for fname in os.listdir(src):
    src_fname = os.path.join(src, fname)
    dst_fname = os.path.join(dst, fname)
    shutil.copy(src_fname, dst_fname)
