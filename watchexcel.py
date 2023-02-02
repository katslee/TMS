import os
import glob
import time
import gen_bulletin

# MacOS development
#watch = "/Users/Kats/Documents/TickerManagementSystem/Python/watch/"

# AMS UAT Server
watch = "/data1/TMS/phrase1/network/export/"

os.chdir(watch)
latest_filename = ""
latest_filemodtime = 0
for file in glob.glob("*.xls*"):
    filemodtime = time.ctime(os.path.getmtime(watch + file))
    if latest_filename == "" or filemodtime > latest_filemodtime:
        latest_filename = file
        latest_filemodtime = filemodtime
gen_bulletin.read_excel(watch + latest_filename)
