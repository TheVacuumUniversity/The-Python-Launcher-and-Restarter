# The-Python-Launcher-and-Restarter
Python code that can launch a program such as excel file and restart the program under some circumstances

I have already achieved to create script that will check for defined errors, in case of error it will kill excel and restart the macro. the problem is that it then waits until the Macro finishes - if another error reoccurs, the CheckForErrors script does nothing.

What I would need is to create a master file which will run two processes at the same time (multiprocessing). One process would be calling the excel macro, the second would check for errors and pontentially kill the excel. It could then send message to the master file that there was an error and the master file would restart the excel macro.

