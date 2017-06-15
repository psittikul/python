import sys
from cx_Freeze import setup, Executable

includefiles = [r"C:\Python36-32\Python Scripts\Mail Merge App\email-icon.ico", r"C:\Python36-32\Python Scripts\Mail Merge App\email-icon.png", r"C:\Python36-32\DLLs\_tkinter.pyd", r"C:\Python36-32\DLLs\tcl86t.dll", r"C:\Python36-32\DLLs\tk86t.dll"]
exclude = ["sqlite3", "html", "http", "json", "adodbapi", "jinja2", "markupsafe"]
build_exe_options = {"packages": ["os", "tkinter", "asyncio"], "include_files":includefiles, "exclude":exclude}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup( name = "mailmerge",
       version = "0.1",
       description = "Orio Mail Merge Application",
       options = {"build_exe": build_exe_options},
       executables = [Executable("testgui.py", base=base, icon="email-icon.ico")])