import os
import glob
import gen_bulletin

# MacOS development
#watch = "/Users/Kats/Documents/TickerManagementSystem/Python/watch/"

# AMS UAT Server
watch = "/data1/TMS/phrase1/network/export/"

os.chdir(watch)
for file in glob.glob("*.xls*"):
    gen_bulletin.read_excel(watch + file)
