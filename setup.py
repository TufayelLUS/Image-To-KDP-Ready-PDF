import sys
from cx_Freeze import setup, Executable

base = None

if sys.platform == "win32":
    base = "Win32GUI"  # Use this option to create a GUI executable on Windows
# set the name of the python script below
executables = [Executable("KDP Format.py", base="Console", icon="icon.ico")]
# if you want to show the console, use the line below instead and comment out the line above
# executables = [Executable("python_file_name.py")]

options = {
    "build_exe": {
        "optimize": 2,
        "packages": [],  # List of packages to include
        "include_files": ["icon.ico", "settings.ini"],  # List of additional files to include
    }
}

setup(
    name="KDP Format Tool", # name of the process
    version="1.0",
    description="KDP Format Tool", # description of the process
    options=options,
    executables=executables
)