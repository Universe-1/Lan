print("{}".format("开发者：兰锦航").center(60,"="))
print("854649049@qq.com".center(67,"="))
print(" ")

import openpyxl as px
import time

namelist=[]
birthdd=[]
showdic={}

def Read(a):
    return  sheet[str(a)].value


wb=px.load_workbook("C:\\Users\\LJH\\Desktop\\srtx\\birthday.xlsx")
sheet=wb.worksheets[0]

date=time.strftime("%m-%d",time.gmtime())
#print(date)
#print(Read("A1"),Read("B34"))

def change(d):
    d=str(d)
    day=0
    day31=[1,3,5,7,8,10,12]
    day30=[4,6,9,11]
    day28=[2]
    if len(d)!=5:
        d=d[5:]
        #print(d)
    months=int(d[0:2])
    for month in range(1,months):
        if month in day31:
            day=day+31
        elif month in day30:
            day=day+30
        elif month in day28:
            day=day+28
    return day+int(d[3:])

def change2(d):
    d=str(d)
    for l in "/_年月日-.":
        d=d.replace(l," ")
        a=d.split()
    #print(a)
    day=0
    day31=[1,3,5,7,8,10,12]
    day30=[4,6,9,11]
    day28=[2]
    if len(a)==3:
        a=a[1:]
    months=int(a[0])
    #print(months)
    for month in range(1,months):
        if month in day31:
            day=day+31
        elif month in day30:
            day=day+30
        elif month in day28:
            day=day+28

    return day+int(a[1])
        
#print(change2("10-8"),"t")
#print(change("2020-10-08"),"t")

def add_name_list():
    rownum=1
    while True:
        name=str(Read("A"+str(rownum)))
        if name=='None':
                 break
        namelist.append(name)
        rownum=rownum+1


def add_day_list():
    rownum=1
    while True:
        day=Read("B"+str(rownum))
        if str(day)=="None":
                break
        day=change2(Read("B"+str(rownum)))
        birthdd.append(day)
        #print(day)
        rownum=rownum+1


add_name_list()
add_day_list()
#print(namelist,birthdd)

pddl=0
date=int(change2(date))
ddl=date+15
if ddl>365:
    pddl=ddl%365
    #print(pddl)
for dex in range(0,len(namelist)):
    if date<=int(birthdd[dex])<=ddl:
        #print(namelist[dex],Read("B"+str(dex+1)),"剩余{}天".format((int(birthdd[dex]-date))))
        showdic[str(namelist[dex])+str(Read("B"+str(dex+1)))]=int(birthdd[dex]-date)
    if 0<=int(birthdd[dex])<=pddl:
         showdic[str(namelist[dex])+str(Read("B"+str(dex+1)))]=int(birthdd[dex]+365-date)
items=list(showdic.items())
items.sort(key=lambda x:x[1])
for k in range(len(items)):
    a,b=items[k]
    print(a,"剩余{}天".format(b))
print("DONE")
time.sleep(15)

