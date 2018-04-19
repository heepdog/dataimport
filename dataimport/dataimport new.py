from ezodf2.conf import config
from ezodf2 import opendoc, Sheet, Table
import easygui
import pandas as pd
import time


filename2 = 'C:/Python34/chris/test.ods'
data = 'data.csv'

files = easygui.fileopenbox(multiple = True)

starttime = time.time()

date = (1,1)
comments = (1,21)
Summarydata = [(2,3),(2,4),(2,5),(2,6),
           (2,7),(2,8),(2,10),(2,11),(2,13),
           (2,14),(2,16),(2,17),(2,19),(2,20)]

fueldata = [(7,43),(7,45)]

avgs1 = [(33,12),(1,1500),(2,1500),(3,1500),(4,1500),   #Train1 CF,Train1 cfm,Train1 dH,Train1 Temp
        (11,1500),(12,1500),(13,1500),(14,1500),        #Amb CF,Amb cfm,Amb dH,Amb Temp
        (15,1500),(16,1500),                            #"Tunnel Temp, Tunnel dP
        (21,1500),(22,1500),                            #"Load Flow,Boiler Flow
        (23,1501),(23,1500),(24,1500),(25,1500),(26,1500),
        (28,1500),(32,1500)]

avgs2 = [(33,12),(1,1000),(2,1000),(3,1000),(4,1000),
        (11,1000),(12,1000),(13,1000),(14,1000),
        (15,1000),(16,1000),
        (21,1000),(22,1000),(23,1001),
        (23,1000),(24,1000),(25,1000),(26,1000),
        (28,1000),(32,1000)]

avgs3 = [(33,12),(1,257),(2,257),(3,257),(4,257),
        (11,257),(12,257),(13,257),(14,257),
        (15,257),(16,257),
        (21,257),(22,257),(23,258),
        (23,257),(24,257),(25,257),(26,257),
        (28,257),(32,257)]

avgs4 = [(33,12),(1,426),(2,426),(3,426),(4,426),
        (11,426),(12,426),(13,426),(14,426),
        (15,426),(16,426),
        (21,426),(22,426),(23,427),
        (23,426),(24,426),(25,426),(26,426),
        (28,426),(32,426)]


labs = [(3,12),(4,12),(5,12),(6,12),    
        (3,13),(4,13),(5,13),(6,13),    
        (3,14),(4,14),(5,14),(6,14),
        (6,16),
        (3,20),(4,20),(5,20),(6,20),
        (3,21),(4,21),(5,21),(6,21),
        (3,22),(4,22),(5,22),(6,22),
        (6,23),
        (3,28),(4,28),(5,28),(6,28),
        (3,29),(4,29),(5,29),(6,29),
        (3,30),(4,30),(5,30),(6,30),
        (6,32)]

#set strategy "all", "all_but_last", "all_less_maxcount'
config.table_expand_strategy.set_strategy('all_less_maxcount',(3000,150))
with open(data,'w') as f:
    #            "Train2 CF,Train2 cfm,Train2 dH,Train2 Temp," +

    f.write("Date,comments,Time(min),Time(hrs),BTU's per hr,High Efficiency, Low Efficiency,Train1 gr/hr,Train1 Lbs/mmBTU,Train2 gr/hr,Train2 Lbs/mmBTU,Amb gr/hr,Amb Lbs/mmBTU,Total gr/hr,Totall Lbs/mmBTU" +
            "Load Weight,Load MC"+
            "Train1 CF,Train1 cfm,Train1 dH,Train1 Temp," +
            "Amb CF,Amb cfm,Amb dH,Amb Temp," +
            "Tunnel Temp, Tunnel dP," +
            "Load Flow,Boiler Flow," +
            "Boiler Out,Boiler In,Load In,Load Out," +
            "Filter Temp,Btu's Captured," +
            "Train1 Front Filter #,Train1 Front Final,Train1 Front Tare,Train1 Front Particulate," +
            "Train1 Rear Filter #,Train1 Rear Final,Train1 Rear Tare,Train1 Rear Particulate," +
            "Train1 Probe Filter #,Train1 Probe Final,Train1 Probe Tare,Train1 Probe Particulate," +
            "Train1 Total," +
            "Train2 Front Filter #,Train2 Front Final,Train2 Front Tare,Train2 Front Particulate," +
            "Train2 Rear Filter #,Train2 Rear Final,Train2 Rear Tare,Train2 Rear Particulate," +
            "Train2 Probe Filter #,Train2 Probe Final,Train2 Probe Tare,Train2 Probe Particulate," +
            "Train2 Total," +
            "Amb Front Filter #,Amb Front Final,Amb Front Tare,Amb Front Particulate," +
            "Amb Rear Filter #,Amb Rear Final,Amb Rear Tare,Amb Rear Particulate," +
            "Amb Probe Filter #,Amb Probe Final,Amb Probe Tare,Amb Probe Particulate," +
            "Amb Total" )

for file in files:
    print("Opening " + file)
    doc = opendoc(file)

    #ss = (doc.sheets._child_by_name('Summary'))
    print("Getting Summary")
    r, c = date
    row =str(doc.sheets._child_by_name('I Data')[c,r].value)
    r, c = comments
    row = row +',"'+str(doc.sheets._child_by_name('I Data')[c,r].value)+'"'

    for i in Summarydata:
        r, c = i
        row = row +', '+ str(doc.sheets._child_by_name('I Data')[c,r].value)
    
    #ss = (doc.sheets._child_by_name('I Fuel'))
    print("Getting Fuel")
    for i in fueldata:
        r, c = i
       # print(ss[c,r].value)
        row = row +', '+ str(doc.sheets._child_by_name('I Fuel')[c,r].value)

    #ss = (doc.sheets._child_by_name('I Data'))
    print("Getting Data")
    if doc.sheets._child_by_name('I Data')[257,0].value == 'Avg/Total':
        avgs= avgs3
    elif doc.sheets._child_by_name('I Data')[425,0].value == 'Avg/Total':
        avgs = avgs4
    elif doc.sheets._child_by_name('I Data')[1000,0].value == 'Avg/Total':
        avgs = avgs2
    else:
        avgs = avgs1


    for i in avgs:
        r, c = i
        #print(ss[c,r].value)
        row = row +', '+ str(doc.sheets._child_by_name('I Data')[c,r].value)

    #ss = (doc.sheets._child_by_name('I Lab'))
    print("Getting Lab")
    for i in labs:
        r, c = i
        #print(ss[c,r].value)
        row = row +', '+ str(doc.sheets._child_by_name('I Lab')[c,r].value)

    with open("data.csv",'a') as f:
        f.write(row + '\n')

print(time.time()-starttime) #took 192 ms