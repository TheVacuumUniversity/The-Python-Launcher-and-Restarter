
# THAT IS WHERE MULTIPROCESSING SHOULD TAKE PLACE I THINK
# What I need is to run the macro and the CheckForErrors script at the same time

# sth like processes = CheckForErrors.py , ExcelMacro
# processes.run

import subprocess

subprocess.run(['Python','Slave2.py'])


#
# def RunMacro():
#     print("Initializing Event macro")
#     xl=win32com.client.Dispatch("Excel.Application")
#     xl.visible = 1
#     xl.Workbooks.Open(os.path.abspath(path.excel_with_macro), ReadOnly=1)
#     xl.Application.Run(path.macro_to_run)
#     # xl.Application.Quit() # Comment this out if your excel script closes
#     del xl
#     print("Event macro ran successfully")