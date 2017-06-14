import sys
import os
os.environ['TCL_LIBRARY'] = "C:\\Python36-32\\tcl\\tcl8.6"
os.environ['TK_LIBRARY'] = "C:\\Python36-32\\tcl\\tk8.6"
from cx_Freeze import setup, Executable


includefiles = [
            (r"C:\Python36-32\Python Scripts\Mail Merge App\email-icon.ico", "email-icon.ico"),
            (r"C:\Python36-32\DLLs\_tkinter.pyd", "_tkinter.pyd"),
            (r"C:\Python36-32\DLLs\tk86t.dll", "tk86t.dll"),
            (r"C:\Python36-32\DLLs\tcl86t.dll", "tcl86t.dll")]
build_exe_options = {"packages": ["os", "asyncio"], "include_files":includefiles}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup( name = "mailmerge",
       version = "0.1",
       description = "Orio Mail Merge Application",
       options = {"build_exe": build_exe_options},
       executables = [Executable("testgui.py", base=base)])
