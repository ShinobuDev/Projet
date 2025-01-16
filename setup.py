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
        ("assets/soliha_logo.png", "assets/soliha_logo.png"),
        ("assets/icon.ico", "assets/icon.ico")
    ]
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Rapprochement Bancaire SOLIHA",
    version="1.0",
    description="Application de rapprochement bancaire pour SOLIHA Normandie",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "main.py",
            base=base,
            icon="assets/icon.ico",
            target_name="RapprochementBancaire_SOLIHA.exe"
        )
    ]
) 