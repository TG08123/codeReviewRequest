#!/usr/bin/env python3.9.2

import os
import PySimpleGUI as sg
import csv
from datetime import datetime, date, timedelta
import datetime as DT
from pptx_template import render
from datetime import datetime
from pytz import reference
import calendar



# Create GUI

sg.theme('Dark Blue 3')  

layout = [[sg.Text('Select PowerPoint template file:')],
            [sg.Input(), sg.FileBrowse(file_types=((".pptx files","*.pptx"),))],
            [sg.Text('Select CSV file containing data:')],
            [sg.Input(), sg.FileBrowse(file_types=((".csv files","*.csv"),))],
            [sg.Text('Select start date for report:')],
            [sg.Input(), sg.CalendarButton('Choose Date')],
            [sg.OK(), sg.Cancel()]]

window = sg.Window('Rapid Report Generator: Calendar Dates to PowerPoint Slide', layout)


# Read and store values from user inputs made via GUI

event, values = window.read()
window.close()
if event == sg.WIN_CLOSED:
    exit()

template_file_path = values[0]
csv_file_path = values[1]
reportStartDate = values[2]



# Define dates based on user-selected start date

date1 = datetime.strptime(reportStartDate, '%Y-%m-%d %H:%M:%S')
date1_string = date1.strftime('%Y-%m-%d')
date1Weekday = calendar.day_name[datetime.date(date1).weekday()]

date2 = date1 + DT.timedelta(days=1)
date2_string = date2.strftime('%Y-%m-%d')
date2Weekday = calendar.day_name[datetime.date(date2).weekday()]

date3 = date1 + DT.timedelta(days=2)
date3_string = date3.strftime('%Y-%m-%d')
date3Weekday = calendar.day_name[datetime.date(date3).weekday()]

date4 = date1 + DT.timedelta(days=3)
date4_string = date4.strftime('%Y-%m-%d')
date4Weekday = calendar.day_name[datetime.date(date4).weekday()]

date5 = date1 + DT.timedelta(days=4)
date5_string = date5.strftime('%Y-%m-%d')
date5Weekday = calendar.day_name[datetime.date(date5).weekday()]

date6 = date1 + DT.timedelta(days=5)
date6_string = date6.strftime('%Y-%m-%d')
date6Weekday = calendar.day_name[datetime.date(date6).weekday()]

date7 = date1 + DT.timedelta(days=6)
date7_string = date7.strftime('%Y-%m-%d')
date7Weekday = calendar.day_name[datetime.date(date7).weekday()]

date8 = date1 + DT.timedelta(days=7)
date8_string = date8.strftime('%Y-%m-%d')

date9 = date1 + DT.timedelta(days=8)
date9_string = date9.strftime('%Y-%m-%d')

date10 = date1 + DT.timedelta(days=9)
date10_string = date10.strftime('%Y-%m-%d')

date11 = date1 + DT.timedelta(days=10)
date11_string = date11.strftime('%Y-%m-%d')

date12 = date1 + DT.timedelta(days=11)
date12_string = date12.strftime('%Y-%m-%d')

date13 = date1 + DT.timedelta(days=12)
date13_string = date13.strftime('%Y-%m-%d')

date14 = date1 + DT.timedelta(days=13)
date14_string = date14.strftime('%Y-%m-%d')

date15 = date1 + DT.timedelta(days=14)
date15_string = date15.strftime('%Y-%m-%d')

date16 = date1 + DT.timedelta(days=15)
date16_string = date16.strftime('%Y-%m-%d')

date17 = date1 + DT.timedelta(days=16)
date17_string = date17.strftime('%Y-%m-%d')

date18 = date1 + DT.timedelta(days=17)
date18_string = date18.strftime('%Y-%m-%d')

date19 = date1 + DT.timedelta(days=18)
date19_string = date19.strftime('%Y-%m-%d')

date20 = date1 + DT.timedelta(days=19)
date20_string = date20.strftime('%Y-%m-%d')

date21 = date1 + DT.timedelta(days=20)
date21_string = date21.strftime('%Y-%m-%d')

date22 = date1 + DT.timedelta(days=21)
date22_string = date22.strftime('%Y-%m-%d')

date23 = date1 + DT.timedelta(days=22)
date23_string = date23.strftime('%Y-%m-%d')

date24 = date1 + DT.timedelta(days=23)
date24_string = date24.strftime('%Y-%m-%d')

date25 = date1 + DT.timedelta(days=24)
date25_string = date25.strftime('%Y-%m-%d')

date26 = date1 + DT.timedelta(days=25)
date26_string = date26.strftime('%Y-%m-%d')

date27 = date1 + DT.timedelta(days=26)
date27_string = date27.strftime('%Y-%m-%d')

date28 = date1 + DT.timedelta(days=27)
date28_string = date28.strftime('%Y-%m-%d')


# Function to get dates and events to fill in Calendar

def getTimePlusEvents(timeDate, event):
    
    timeDate_splitT = timeDate.split(" ")
    time = timeDate_splitT[1] 

    fromDate = datetime.strptime(row0_time, '%Y-%m-%d')
    toDate = datetime.strptime(row1_time, '%Y-%m-%d')
    delta = toDate - fromDate
    timeDiffDays = delta.days 

    if timeDiffDays != 0 or event.lower() == "TRAINING HOLIDAY".lower():
        date_timePlusEvents = '*' + event.upper() + '*' + '\n' 

    else:
        date_timePlusEvents = time + ' - ' + event + '\n'  

    return date_timePlusEvents



# Import user-selected CSV file, parse and store data based on date

with open(csv_file_path) as csv_file:
    data = csv.reader(csv_file, delimiter=',')
    linenum = 0
    
    date1_timePlusEvents = ' '
    date2_timePlusEvents = ' '
    date3_timePlusEvents = ' '
    date4_timePlusEvents = ' '
    date5_timePlusEvents = ' '
    date6_timePlusEvents = ' '
    date7_timePlusEvents = ' '
    date8_timePlusEvents = ' '
    date9_timePlusEvents = ' '
    date10_timePlusEvents = ' '
    date11_timePlusEvents = ' '
    date12_timePlusEvents = ' '
    date13_timePlusEvents = ' '
    date14_timePlusEvents = ' '
    date15_timePlusEvents = ' '
    date16_timePlusEvents = ' '
    date17_timePlusEvents = ' '
    date18_timePlusEvents = ' '
    date19_timePlusEvents = ' '
    date20_timePlusEvents = ' '
    date21_timePlusEvents = ' '
    date22_timePlusEvents = ' '
    date23_timePlusEvents = ' '
    date24_timePlusEvents = ' '
    date25_timePlusEvents = ' '
    date26_timePlusEvents = ' '
    date27_timePlusEvents = ' '
    date28_timePlusEvents = ' '

    for row in data:

        if linenum == 0: 

            linenum += 1

        elif linenum != 0:
            row0_timeA = datetime.strptime(row[0], '%m/%d/%Y %H:%M')
            row1_timeA = datetime.strptime(row[1], '%m/%d/%Y %H:%M')
            row0_time = row0_timeA.strftime('%Y-%m-%d')
            row1_time = row1_timeA.strftime('%Y-%m-%d')
        
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date1_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date1_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date2_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date2_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date3_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date3_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date4_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date4_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date5_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date5_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date6_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date6_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date7_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date7_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date8_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date8_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date9_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date9_timePlusEvents += getTimePlusEvents(row[0], row[2])
     
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date10_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date10_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date11_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date11_timePlusEvents += getTimePlusEvents(row[0], row[2])   

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date12_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date12_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date13_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date13_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date14_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date14_timePlusEvents += getTimePlusEvents(row[0], row[2])
           
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date15_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date15_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date16_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date16_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date17_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date17_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date18_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date18_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date19_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date19_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date20_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date20_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date21_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date21_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date22_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date22_timePlusEvents += getTimePlusEvents(row[0], row[2])
            
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date23_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date23_timePlusEvents += getTimePlusEvents(row[0], row[2])
  
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date24_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date24_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date25_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date25_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date26_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date26_timePlusEvents += getTimePlusEvents(row[0], row[2])
  
            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date27_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date27_timePlusEvents += getTimePlusEvents(row[0], row[2])

            if datetime.strptime(row0_time, '%Y-%m-%d') <= datetime.strptime(date28_string, '%Y-%m-%d') <= datetime.strptime(row1_time, '%Y-%m-%d'):
                date28_timePlusEvents += getTimePlusEvents(row[0], row[2])
  

        else:     

            linenum += 1


            
# Reformatting data to go into PowerPoint template

date1_label = date1.strftime('%m.%d')
date2_label = date2.strftime('%m.%d')
date3_label = date3.strftime('%m.%d')
date4_label = date4.strftime('%m.%d')
date5_label = date5.strftime('%m.%d')
date6_label = date6.strftime('%m.%d')
date7_label = date7.strftime('%m.%d')
date8_label = date8.strftime('%m.%d')
date9_label = date9.strftime('%m.%d')
date10_label = date10.strftime('%m.%d')
date11_label = date11.strftime('%m.%d')
date12_label = date12.strftime('%m.%d')
date13_label = date13.strftime('%m.%d')
date14_label = date14.strftime('%m.%d')
date15_label = date15.strftime('%m.%d')
date16_label = date16.strftime('%m.%d')
date17_label = date17.strftime('%m.%d')
date18_label = date18.strftime('%m.%d')
date19_label = date19.strftime('%m.%d')
date20_label = date20.strftime('%m.%d')
date21_label = date21.strftime('%m.%d')
date22_label = date22.strftime('%m.%d')
date23_label = date23.strftime('%m.%d')
date24_label = date24.strftime('%m.%d')
date25_label = date25.strftime('%m.%d')
date26_label = date26.strftime('%m.%d')
date27_label = date27.strftime('%m.%d')
date28_label = date28.strftime('%m.%d')


# Populate PowerPoint template

model = {
    
    "date1Weekday":date1Weekday,
    "date2Weekday":date2Weekday,
    "date3Weekday":date3Weekday,
    "date4Weekday":date4Weekday,
    "date5Weekday":date5Weekday,
    "date6Weekday":date6Weekday,
    "date7Weekday":date7Weekday,
    
    "date1": date1_label,
    "event1": date1_timePlusEvents,
    "date2": date2_label,
    "event2": date2_timePlusEvents,
    "date3": date3_label,
    "event3": date3_timePlusEvents,
    "date4": date4_label,
    "event4": date4_timePlusEvents,
    "date5": date5_label,
    "event5": date5_timePlusEvents,
    "date6": date6_label,
    "event6": date6_timePlusEvents,
    "date7": date7_label,
    "event7": date7_timePlusEvents,
    "date8": date8_label,
    "event8": date8_timePlusEvents,
    "date9": date9_label,
    "event9": date9_timePlusEvents,
    "date10": date10_label,
    "event10": date10_timePlusEvents,
    "date11": date11_label,
    "event11": date11_timePlusEvents,
    "date12": date12_label,
    "event12": date12_timePlusEvents,
    "date13": date13_label,
    "event13": date13_timePlusEvents,
    "date14": date14_label,
    "event14": date14_timePlusEvents,
    "date15": date15_label,
    "event15": date15_timePlusEvents,
    "date16": date16_label,
    "event16": date16_timePlusEvents,
    "date17": date17_label,
    "event17": date17_timePlusEvents,
    "date18": date18_label,
    "event18": date18_timePlusEvents,
    "date19": date19_label,
    "event19": date19_timePlusEvents,
    "date20": date20_label,
    "event20": date20_timePlusEvents,
    "date21": date21_label,
    "event21": date21_timePlusEvents,
    "date22": date22_label,
    "event22": date22_timePlusEvents,
    "date23": date23_label,
    "event23": date23_timePlusEvents,
    "date24": date24_label,
    "event24": date24_timePlusEvents,
    "date25": date25_label,
    "event25": date25_timePlusEvents,
    "date26": date26_label,
    "event26": date26_timePlusEvents,
    "date27": date27_label,
    "event27": date27_timePlusEvents,
    "date28": date28_label,
    "event28": date28_timePlusEvents
}



# Get current date, time, timezone for automatic file naming convention

now = datetime.now()  
localtime = reference.LocalTimezone()  
localtime.tzname(now)
DateForFile = now.strftime('%Y%m%d_%H%M%S')  


# Find file directory path (without filename) to save new filename in same place

fileDirectory = os.path.dirname(template_file_path)


# Send file to output path

inputPath = template_file_path
outputPath = fileDirectory + "/" + DateForFile + '_Report.pptx'
render.render_pptx(inputPath, model, outputPath)


# User Dialog, open final document

sg.popup('\nDocument successfully generated!')
os.system(outputPath)
