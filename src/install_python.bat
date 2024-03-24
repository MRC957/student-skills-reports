:: @echo off

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed. Installing Python3...

    :: Download Python installer
    bitsadmin.exe /transfer "PythonInstaller" https://www.python.org/ftp/python/3.12.1/python-3.12.1-amd64.exe "%cd%\python-3.12.1-amd64.exe"

    :: Install Python3
    .\python-3.12.1-amd64.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

    echo Python installation is finished. The computer will be restarted...
    shutdown /r
)
