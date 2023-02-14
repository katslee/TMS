import gen_ordering
import os
import glob
import filecmp
import shutil
from datetime import datetime

# MacOS Development
#pythonfolder = "/Users/Kats/Documents/TickerManagementSystem/Python/python/"
#updatefolder = "/Users/Kats/Documents/TickerManagementSystem/Python/update/"
#text_output = "/Users/Kats/Documents/TickerManagementSystem/Python/TextBulletin/"
#graphic_output = "/Users/Kats/Documents/TickerManagementSystem/Python/GraphicBulletin/"
#working = "/Users/Kats/Documents/TickerManagementSystem/Python/working/"

# AMS UAT Server
pythonfolder = "/data1/TMS/phrase1/python/"
updatefolder = "/data1/TMS/phrase1/update/"
text_output = "/data1/TMS/phrase1/network/export/result/TextBulletin/"
graphic_output = "/data1/TMS/phrase1/network/export/result/GraphicBulletin/"
graphicengine1 = "/data1/TMS/phrase1/network/graphicengine1/result/GraphicBulletin/"
graphicengine2 = "/data1/TMS/phrase1/local/graphicengine2/result/GraphicBulletin/"
textengine1 = "/data1/TMS/phrase1/network/graphicengine1/result/TextBulletin/"
textengine2 = "/data1/TMS/phrase1/local/graphicengine2/result/TextBulletin/"
working = "/data1/TMS/phrase1/working/"

def comparefiles(f1, f2):
    result = filecmp.cmp(f1,f2,shallow=False)
    return result

os.chdir(updatefolder)
for file in glob.glob("*.xls*"):
    gen_ordering.gen_order(updatefolder + file, working, working)

    # copy order file if different
    g_result = comparefiles(working + "gb_order.txt", updatefolder + "gb_order.txt")
    if not g_result:
        try:
            shutil.copy2(working + "gb_order.txt", graphic_output + "gb_order.txt")
        except:
            print("Error.")

        try:
            shutil.copy2(working + "gb_order.txt", updatefolder + "gb_order.txt")
        except:
            print("Error.")

        try:
            shutil.copy2(working + "gb_order.txt", graphicengine1 + "gb_order.txt")
        except:
            print("Error.")

        try:
            shutil.copy2(working + "gb_order.txt", graphicengine2 + "gb_order.txt")
        except:
            print("Error.")

    t_result = comparefiles(working + "L-Title.txt", updatefolder + "L-Title.txt")
    if not t_result:
        try:
            shutil.copy2(working + "L-Title.txt", text_output + "L-Title.txt")
        except:
            print("Error.")

        try:
            shutil.copy2(working + "L-Title.txt", updatefolder + "L-Title.txt")
        except:
            print("Error.")

        try:
            shutil.copy2(working + "L-Title.txt", textengine1 + "L-Title.txt")
        except:
            print("Error.")

        try:
            shutil.copy2(working + "L-Title.txt", textengine2 + "L-Title.txt")
        except:
            print("Error.")
        