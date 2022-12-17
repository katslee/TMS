import openpyxl
from datetime import datetime, date, timedelta
import os
from subprocess import call
import glob
import shutil
import logging
import gen_ordering

crlf = chr(13) + chr(10)
lf = chr(10)
output = "/Users/Kats/Documents/TickerManagementSystem/Python/"
working = "/Users/Kats/Documents/TickerManagementSystem/Python/working/"
text_bulletin_filename = "L-Title.txt"

def read_excel(filename):
    global output
    global working
    global lf
    global crlf
    global text_bulletin_filename

    #folder = datetime.now().strftime("%d%m%Y%H%M%S") + "/"
    folder = datetime.now().strftime("%Y%m%d%H%M") + "/"
    os.mkdir(output + folder)
    files = glob.glob(working + "*")
    for f in files:
            os.remove(f)
    logging.basicConfig(filename=output + folder + "error_" + os.path.basename(filename) + ".txt",level=logging.ERROR)
    error = False
    wb = openpyxl.load_workbook(filename, data_only=True)
    ws = wb.worksheets[0]

    script = working + "genbulletin.sh"
    with open(script, "w") as f:
        f.writelines("# " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") +"\n")
        f.writelines("cd " + working + "\n")

    row = 2
    while ws.cell(row=row,column=1).value != None:
        bulletinType = str.lower(ws.cell(row=row, column=1).value)
        if bulletinType[0:4] == "text":
            bulletinType = "T"
        else:
            bulletinType = "G"
        sn = str(ws.cell(row=row, column=11).value)
        title = ws.cell(row=row, column=12).value

        if title.count(lf) > 1:
            logging.error(sn + " - title line exceed (" + str(title.count(lf)) + ").\n")
            error = True
            for line in title.splitlines():
                if len(line) > 11:
                    logging.error(sn +  " - title words exceed (" + str(len(line)) + ").\n")
                    error = True

        content = ws.cell(row=row, column=13).value
        if content.count(lf) > 3:
            logging.error(sn + " - content line exceed (" + str(content.count(lf)) + ").\n")
            error = True
        for line in content.splitlines():
                if len(line) > 6:
                    logging.error(sn + " - content words exceed (" + str(len(line)) + ").\n")
                    error = True
        content = content.replace(lf, crlf)

        footer = ws.cell(row=row, column=14).value
        if footer == None:
            footer = " "
        if footer.count(lf) > 1:
            logging.error(sn + " - footer line exceed ("  + str(footer.count(lf)) + ").\n")
            error = True
        for line in footer.splitlines():
                if len(line) > 10:
                    logging.error(sn + " - footer words exceed (" + str(len(line)) + ").\n")
                    error = True

        qrcode = ws.cell(row=row, column=15).value

        txdate = ws.cell(row=row, column=5).value
        enddate = ws.cell(row=row, column=7).value
        if enddate == None:
            enddate = txdate + timedelta(days=14)
        etime = ws.cell(row=row, column=8).value
        if etime == None:
            endtime = timedelta(hours=23, minutes=59)
        else:
            endtime = timedelta(hours=int(etime / 100), minutes=etime % 100)
        EndTime = endtime + enddate

        fname = working + "title" + sn + ".txt"
        with open(fname, "w") as f:
            f.writelines(title)

        fname = working  + "content" + sn + ".txt"
        with open(fname, "w") as f:
            f.writelines(content)

        fname = working + "footer" + sn + ".txt"
        with open(fname, "w") as f:
            f.writelines(str(footer))
        with open(script, "a") as f:
            if EndTime.month < 10:
                dcode = "0" + str(EndTime.month)
            else:
                dcode = str(EndTime.month)
            if EndTime.day < 10:
                dcode = dcode + "0" + str(EndTime.day)
            else:
                dcode = dcode + str(EndTime.day)
            if EndTime.hour < 10:
                tcode = "0" + str(EndTime.hour)
            else:
                tcode = str(EndTime.hour)
            if EndTime.minute < 10:
                tcode = tcode + "0" + str(EndTime.minute)
            else:
                tcode = tcode + str(EndTime.minute)
            if bulletinType == "G":
                f.writelines("../upper_image_billboard.sh " + str(sn) + " " + chr(34) + qrcode + chr(34) + " " + dcode + " " + tcode + "\n")
            else:
                with open(working + gen_ordering.fname(sn, EndTime, "T"), "w") as f:
                    f.writelines(title + crlf)
                    f.writelines(content + crlf)
                    f.writelines(crlf)

            row += 1

    os.chmod(script,0o755)
#    if error or not error:
#        rcode = call(script,shell=True)
#        files = glob.iglob(os.path.join(working, "*.jpg"))
#        for f in files:
#            if os.path.isfile(f):
#                shutil.copy2(f, output + folder)



