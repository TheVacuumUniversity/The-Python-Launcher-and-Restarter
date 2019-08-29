# Master script that will run the slave {Refresher} as a subprocess
# At the same time it will check for possible errors
# If an error occurs, it will restart the subprocess

import subprocess
from Slave import KillSlave


def stringcompare(looking_for_windows, list_of_active_windows):
    result = 0
    for lwindow in looking_for_windows:
        for window in list_of_active_windows:
            if str(window.find(lwindow)) != '-1':
                result =+ 1
    return result

def check_for_errors():
    # Check for errors (empty excel opens; BW returns error....)
    while True:
        errmsg = ['Book1', 'BW Sucks', 'Bex Analyzer Message Window', 'BW Server Error', 'Save as',
                  'Microsoft Visual Basic']
        window_list = gw.getAllTitles()
        # print(window_list)
        print("Check for errors process ID", multiprocessing.current_process().pid)
        if stringcompare(errmsg, window_list) > 0:
            print("Error occured, restarting the Slave")
            KillSlave()
#             should also kill excel and access

#  check for error does nothing :(((
while True:
    check_for_errors()
    subprocess.run(["Python", "Slave.py"])
    check_for_errors()