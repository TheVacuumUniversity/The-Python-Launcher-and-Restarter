from config import path
import win32com.client
import datetime
import os.path
import psutil


def RunMacro():
    print("Master: Initializing Event macro")
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Visible = True
    xl.Workbooks.Open(os.path.abspath(path.excel_with_macro), ReadOnly=1)
    xl.Application.Run(path.macro_to_run)
    # xl.Application.Quit() # Comment this out if your excel script closes
    del xl
    # print("Master: Event macro initialized successfully")

def kill_process_function(procname):
    for proc in psutil.process_iter():
        # check whether the process name matches
        if proc.name() == procname:
            proc.kill()

while True:
    # wait for time
    a = datetime.datetime.now().time()
    if a.hour == 6 and a.minute == 0 and a.second == 0:
        print("Master: time to initiate Event macro", a.hour,a.minute,a.second)
        kill_process_function("EXCEL.EXE")
        kill_process_function("ACCESS.EXE")
        RunMacro()

