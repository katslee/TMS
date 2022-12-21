import os
import glob
import gen_bulletin

basefolder = "/data1/TMS/phrase1/"
watch = basefolder + "user/ingest/"

os.chdir(watch)
for file in glob.glob("*.xls*"):
    gen_bulletin.read_excel(watch + file)
