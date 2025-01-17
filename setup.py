import sys
from cx_Freeze import setup, Executable

# Dépendances
build_exe_options = {
    "packages": [
        "tkinter", 
        "pandas", 
        "openpyxl", 
        "pathlib", 
        "customtkinter", 
        "PIL"
    ],
    "excludes": [],
    "include_files": [
        ("assets/_logo.png", "assets/_logo.png"),
        ("assets/icon.ico", "assets/icon.ico")
    ]
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="RA-PRO",
    version="1.0",
    description="Application de rapprochement bancaire",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "main.py",
            base=base,
            icon="assets/icon.ico",
            target_name="RA-PRO.exe"
        )
    ]
) 