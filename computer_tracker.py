import pandas as pd
import numpy as np
from pandas import ExcelFile, ExcelWriter
import time
import re
import datetime
import json
from datetime import timedelta
import win32gui
import win32con
from win32gui import GetWindowText, GetForegroundWindow


class activity:

    def __init__(self,name,start):
        self.name = name
        self.start = start
        self.state = 'not active'

    def get_time(self):
        self.current = datetime.datetime.now()
        return self.current

    def duration(self):
        self.get_time()
        duration = self.current-self.start
        if duration.seconds > 60:
            hours = duration.seconds // 3600
            minutes = duration.seconds // 60
            seconds = duration.seconds % 60
        else:
            hours = 0
            minutes = 0
            seconds = duration.seconds
        self.entry = {'Application':[self.name],'Hours':[hours],'Minutes':[minutes],'Seconds':[seconds],}
        return duration, self.entry

    def log_to_excel(self):
        self.duration()
        complete_logs.append(pd.DataFrame(self.entry))
        excel_df = pd.concat(complete_logs,ignore_index=True,sort=False)

        with pd.ExcelWriter(file) as writer:
            excel_df.to_excel(writer,sheet_name='Activities')
            writer.save()

file='logs.xlsx'
complete_logs = []

try:
    logs = pd.read_excel(file,sheet_name='Activities',index_col=0)
    complete_logs.append(logs)
except FileNotFoundError:
    logs = pd.DataFrame({'Application':[],'Hours':[],'Minutes':[],'Seconds':[],})
    with pd.ExcelWriter(file) as writer:
        logs.to_excel(writer,sheet_name='Activities')

active_window = activity(None,0)
new_window = activity(None,0)

while True:
    try:
        #get name of active window
        window = GetForegroundWindow()
        new_window.name = GetWindowText(GetForegroundWindow())

        #change the name of some active windows
        if re.search('Atom',new_window.name):
            new_window.name = 'Atom'
        elif re.search('Google Chrome',new_window.name):
            new_window.name = 'Chrome'

            #check if window was already active
        if active_window.name != new_window.name:

                #if user opens new window
            if active_window.state == 'active':
                active_window.log_to_excel()

            print('You switched to %s.' % new_window.name)
            active_window.name = new_window.name
            active_window.start = datetime.datetime.now()

        else:
            #if window was active display duration
            active_window.state = 'active'

        new_window.name=active_window.name
        time.sleep(5)

    except KeyboardInterrupt:
            complete_logs.append(pd.DataFrame(active_window.entry))
            excel_df = pd.concat(complete_logs,ignore_index=True,sort=False)

            with pd.ExcelWriter(file) as writer:
                excel_df.to_excel(writer,sheet_name='Activities')
                writer.save()
