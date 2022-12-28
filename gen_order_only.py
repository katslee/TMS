import gen_ordering
import os
import glob
import filecmp
import shutil
from datetime import datetime

#watch = "/Users/Kats/Documents/TickerManagementSystem/Python/"
#watch = "/data1/TMS/phrase1/user/ingest/"
#output = "/data1/TMS/phrase1/user/result/"
#working = "/data1/TMS/phrase1/working/"
pythonfolder = "/data1/TMS/phrase1/python/"
updatefolder = "/data1/TMS/phrase1/update/"

#watch = "/data1/TMS/phrase1/network/export/"
text_output = "/data1/TMS/phrase1/network/export/result/TextBulletin/"
graphic_output = "/data1/TMS/phrase1/network/export/result/GraphicBulletin/"
working = "/data1/TMS/phrase1/working/"

def comparefiles(f1, f2):
    result = filecmp.cmp(f1,f2,shallow=False)
    return result

os.chdir(updatefolder)
for file in glob.glob("*.xls*"):
    #folder = datetime.now().strftime("%Y%m%d%H%M") + "/"
    #os.mkdir(output + folder)
    #gen_ordering.gen_order(updatefolder + file, output + folder)
    gen_ordering.gen_order(updatefolder + file, working, working)

    # copy order file if different
    g_result = comparefiles(working + "gb_order.txt", updatefolder + "gb_order.txt")
    if not g_result:
        shutil.copy2(working + "gb_order.txt", graphic_output + "gb_order.txt")
        shutil.copy2(working + "gb_order.txt", updatefolder + "gb_order.txt")

    t_result = comparefiles(working + "L-Title.txt", updatefolder + "L-Title.txt")
    if not t_result:
        shutil.copy2(working + "L-Title.txt", text_output + "L-Title.txt")
        shutil.copy2(working + "L-Title.txt", updatefolder + "L-Title.txt")