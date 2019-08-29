# Script to call Mail macro from Scheduller
import os.path
import time
import win32com.client
import psutil
import pygetwindow as gw
import multiprocessing
from multiprocessing import Process


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

    # xl.Workbooks.Open(os.path.abspath("\\\\SECZEFNPBRN003\\BRNO FSC\\BI\\AdvancedEmailScheduler\\MainExcel_Development_v2.2.xlam"), ReadOnly=1)
    # xl.Application.Run("'\\\\SECZEFNPBRN003\\BRNO FSC\\BI\\AdvancedEmailScheduler\\MainExcel_Development_v2.2.xlam'!Timer.EventMacro")

    xl.Workbooks.Open(os.path.abspath("\\\\SECZEFNPBRN003\\BRNO FSC\\BI\\PythonEmailScheduler\\testingwb.xlsm"), ReadOnly=1)
    xl.Application.Run("'\\\\SECZEFNPBRN003\\BRNO FSC\\BI\\PythonEmailScheduler\\testingwb.xlsm'!Module1.test")

    xl.Application.Quit() # Comment this out if your excel script closes
    del xl
    print("Event macro ran successfully")

def stringcompare(looking_for_windows, list_of_active_windows):
    result = 0
    for lwindow in looking_for_windows:
        for window in list_of_active_windows:
            if str(window.find(lwindow)) != '-1':
                result =+ 1
    return result


def run_subprocess():  #bezi jedno po druhem, ne zaroven

    processes = [Process(target=RunMacro()), Process(target=check_for_errors())]

    for p in processes:
        p.name = multiprocessing.current_process().name
        print(p.name , p.pid)
        p.start()

    for p in processes:
        p.join()

    return processes

def kill_subprocess(processes):
    for p in processes:
        p.terminate()

def check_for_errors():
    while True:
        errmsg = ['Book1', 'BW Sucks', 'Bex Analyzer Message Window', 'BW Server Error', 'Save as',
                  'Microsoft Visual Basic']
        window_list = gw.getAllTitles()
        # print(window_list)
        # print("Check for errors process ID", multiprocessing.current_process().pid)
        if stringcompare(errmsg, window_list) > 0:
            print("Error occured, restarting the Event macro")
            kill_subprocess(running_process)  #try to get process ID
            running_process = run_subprocess()  #nespusti znovu


def main():

    running_process = run_subprocess()


main()