import openpyxl
from datetime import datetime, date, timedelta

watch = "/home/kats/TMS/"
def takefrequency(elem):
    return elem[1]

def takepriority(elem):
    return elem[2]

def fname(sn, etime, btype):
    if etime.month < 10:
        dcode = "0" + str(etime.month)
    else:
        dcode = str(etime.month)
    if etime.day < 10:
        dcode = dcode + "0" + str(etime.day)
    else:
        dcode = dcode + str(etime.day)
    if etime.hour < 10:
        tcode = "0" + str(etime.hour)
    else:
        tcode = str(etime.hour)
    if etime.minute < 10:
        tcode = tcode + "0" + str(etime.minute)
    else:
        tcode = tcode + str(etime.minute)
    if btype == "G":
        fn = sn + " d "  + dcode + " " + tcode + ".jpg"
    else:
        fn = sn + " d "  + dcode + " " + tcode + ".txt"
    return fn

def remove_dup(bulletins):
    last_bulletin = []
    new_bulletins = []
    for b in bulletins:
        if last_bulletin != b:
            new_bulletins.append(b)
        last_bulletin = b
    return new_bulletins

def gen_order(filename):
# Bulletin [sn, bulletinType, frequency, priority, channel, TXTime, EndTime, Filename]

    wb = openpyxl.load_workbook(filename,data_only=True)
    ws = wb.worksheets[0]

    row = 2
    g_bulletins = []
    t_bulletins = []
    while ws.cell(row=row,column=1).value != None:
        sn = str(ws.cell(row=row, column=11).value)
        bulletinType = str.lower(ws.cell(row=row, column=1).value)
        if bulletinType[0:4] == "text":
            bulletinType = "T"
        else:
            bulletinType = "G"
        frequency = ws.cell(row=row, column=2).value
        priority = ws.cell(row=row, column=3).value
        channel = ws.cell(row=row, column=4).value
        txdate = ws.cell(row=row, column=5).value
        txtime = ws.cell(row=row, column=6).value
        if txtime == None:
            txtime = timedelta(hours=0,minutes=0)
        TXTime = txdate + txtime
        enddate = ws.cell(row=row, column=7).value
        if enddate == None:
            enddate  = txdate + timedelta(days=14)
        etime = ws.cell(row=row, column=8).value
        if etime == None:
            endtime = timedelta(hours=23,minutes=59)
        else:
            endtime = timedelta(hours=int(etime/100),minutes=etime % 100)
        EndTime = endtime + enddate
        bulletin = []
        bulletin.append(sn)
        bulletin.append(frequency)
        bulletin.append(priority)
        bulletin.append(channel)
        bulletin.append(TXTime)
        bulletin.append(EndTime)
        bulletin.append(fname(sn,EndTime,bulletinType))
        if bulletinType == "T":
            t_bulletins.append(bulletin)
        else:
            g_bulletins.append(bulletin)
        row += 1

    wb.close

# Graphic Bulletin Order (Priority > Frequency)
    g_bulletins.sort(key=takepriority, reverse=True)
    g_order = []
    finish = False
    while not finish:
        finish = True
        for bulletin in g_bulletins:
            current_time = datetime.now()
            if (bulletin[2] > 0) and (bulletin[4] < current_time) and (current_time <= bulletin[5]):
                #g_order.append(bulletin[6] + "," + str(bulletin[2]))
                g_order.append(bulletin[6])
                bulletin[2] -= 1
                finish = False

# Text Bulletin Order (Priority > Frequency)
    t_bulletins.sort(key=takepriority, reverse=True)
    t_order = []
    finish = False
    while not finish:
        finish = True
        for bulletin in t_bulletins:
            current_time = datetime.now()
            if (bulletin[2] > 0) and (bulletin[4] < current_time) and (current_time <= bulletin[5]):
                g_order.append(bulletin[6] + "," + str(bulletin[2]))
                bulletin[2] -= 1
                finish = False

    g_order = remove_dup(g_order)
    t_order = remove_dup(t_order)

    return  g_order, t_order

graphic_order, text_order = gen_order(watch + "TMS_real case_20221207_v2.xlsx")
print("Graphic")
for b in graphic_order:
    print(b)

print("Text")
for b in text_order:
    print(b)
