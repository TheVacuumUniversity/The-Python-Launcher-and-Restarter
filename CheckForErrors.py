import ctypes
import psutil
import os.path
import time
import win32com.client
from config import path
import datetime

#Run the Timer.Event macro
def RunMacro():
    print("Initializing Event macro")
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Visible = True
    xl.Workbooks.Open(os.path.abspath(path.excel_with_macro), ReadOnly=1)
    xl.Application.Run(path.macro_to_run)
    # xl.Application.Quit() # Comment this out if your excel script closes
    del xl
    print("Event macro ran successfully")

#Function to kill a process (used for Excel and Access)
def kill_process_function(procname):
    for proc in psutil.process_iter():
        # check whether the process name matches
        if proc.name() == procname:
            proc.kill()

# Function checking all the active windows
EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
def foreach_window(hwnd, lParam):
    if IsWindowVisible(hwnd):
        length = GetWindowTextLength(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        GetWindowText(hwnd, buff, length + 1)
        titles.append(buff.value)
    return True

# Function checking if one of defined non-desired windows (errmsg list) is active
# If yes, it will kill excell and access
errmsg = ['Book1 - Excel', 'BW Sucks', 'Bex Analyzer Message Window', 'BW Server Error', 'Save as', 'Microsoft Visual Basic']
def check_result():
    error_occured = 0
    for err in errmsg:
        if err in titles:
            error_occured += 1
            print("Error occured. Err type: ", err)
            kill_process_function("EXCEL.EXE")
            kill_process_function("MSACCESS.EXE")
            RunMacro()

        else:
            error_occured += 0
            # print("no error")
        return error_occured
    return  error_occured

# Run the macro and check for errors at the same time
while True:
    titles = []
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    check_result()



