@echo off
setlocal

set PYTHON_VERSION=3.11.3
set PYTHON_INSTALLER=python-%PYTHON_VERSION%-amd64.exe

echo Checking if Python is already installed...
python --version > nul 2>&1
if %errorlevel% equ 0 (
    echo Python is already installed.
) else (
    echo Python is not installed. Downloading installer...
    curl -o %PYTHON_INSTALLER% https://www.python.org/ftp/python/%PYTHON_VERSION%/%PYTHON_INSTALLER%
    echo Installing Python...
    %PYTHON_INSTALLER% /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    echo Python installation complete.
)

echo "Python installation complete."

echo "Installing required Python modules..."
pip install requests
pip install openpyxl
echo "Installation Complete.
