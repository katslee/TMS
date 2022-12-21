import gen_ordering
import os
import glob
from datetime import datetime

#watch = "/Users/Kats/Documents/TickerManagementSystem/Python/"
watch = "/data1/TMS/phrase1/user/ingest/"
output = "/data1/TMS/phrase1/user/result/"
working = "/data1/TMS/phrase1/working/"
pythonfolder = "/data1/TMS/phrase1/python/"
updatefolder = "/data1/TMS/phrase1/update/"

os.chdir(updatefolder)
for file in glob.glob("*.xls*"):
    folder = datetime.now().strftime("%Y%m%d%H%M") + "/"
    os.mkdir(output + folder)
    gen_ordering.gen_order(updatefolder + file, output + folder)