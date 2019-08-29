# Slave script which will be run as a subprocess
# It should kill access and excel and run the refreshing macro

import os.path
import time
import win32com.client
import psutil
import multiprocessing
import sys
from config import path

#Kill Excel and Access
def kill_process_function(procname):
    for proc in psutil.process_iter():
        # check whether the process name matches
        if proc.name() == procname:
            proc.kill()

#Run the Timer.Event macro
def RunMacro():
    print("Killing Access and Excell")
    print("run macro ID", multiprocessing.current_process().pid)
    kill_process_function("EXCEL.EXE")
    kill_process_function("MSACCESS.EXE")
    print("Excel and Access killed")
    print("Initializing Event macro")
    xl=win32com.client.Dispatch("Excel.Application")
    xl.visible = 1
    xl.Workbooks.Open(os.path.abspath(path.excel_with_macro), ReadOnly=1)
    xl.Application.Run(path.macro_to_run)

    xl.Application.Quit() # Comment this out if your excel script closes
    del xl
    print("Event macro ran successfully")

# Kill the slave in case of error via Master.py
def KillSlave():
    sys.exit()

# If no error, run in loop
while True:
    RunMacro()
    time.sleep(60)

