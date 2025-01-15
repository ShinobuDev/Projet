import sys
from cx_Freeze import setup, Executable

# Dépendances
build_exe_options = {
    "packages": ["tkinter", "pandas", "openpyxl", "pathlib"],
    "excludes": [],
    "include_files": []
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Rapprochement Bancaire",
    version="1.0",
    description="Application de rapprochement bancaire",
    options={"build_exe": build_exe_options},
    executables=[Executable("bank_reconciliation.py", base=base)]
) 