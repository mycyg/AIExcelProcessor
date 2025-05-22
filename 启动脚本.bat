@echo off
:: Set codepage to UTF-8. This might help if any non-ASCII paths are involved.
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ================================================================================
echo Excel Batch Processing Tool - Environment Setup and Startup Script
echo Please run this script as Administrator to ensure Python and libraries can be installed correctly.
echo ================================================================================
echo.

:: Define application main file and requirements file
set APP_FILE=qt_app.py
set REQUIREMENTS_FILE=requirements.txt
set PYTHON_INSTALLER_URL=https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe
set PYTHON_INSTALLER_NAME=python_installer.exe
set PYTHON_COMMAND=python

:: Check Python environment (prioritize py.exe)
echo Checking Python environment...

:: Try using py.exe
py --version > nul 2>&1
if %errorlevel% equ 0 goto py_found
goto py_not_found

:py_found
    echo Python Launcher (py.exe) found. Will use py.exe.
    set PYTHON_COMMAND=py
    goto python_env_checked

:py_not_found
    echo Python Launcher (py.exe) not found. Trying python.exe...
    python --version > nul 2>&1
    if %errorlevel% equ 0 goto python_exe_found
    goto python_needs_install

:python_exe_found
    echo Python (python.exe) found.
    set PYTHON_COMMAND=python
    goto python_env_checked

:python_needs_install
    echo Python (python.exe) is not installed or not in PATH.
    echo Attempting to download and install Python 3.11.5 (64-bit).
    
    :: Check if curl is available
    curl --version > nul 2>&1
    if %errorlevel% neq 0 (
        echo curl command not found. Please install Python 3.11 or higher manually and add to PATH.
        echo Or ensure curl.exe is available (usually in System32).
        pause
        exit /b 1
    )

    echo Downloading Python installer...
    curl -L -o %PYTHON_INSTALLER_NAME% %PYTHON_INSTALLER_URL%
    if %errorlevel% neq 0 (
        echo Failed to download Python installer. Check network or download manually.
        echo URL: %PYTHON_INSTALLER_URL%
        pause
        exit /b 1
    )
    
    echo Installing Python silently (This may take a few minutes and require admin rights).
    echo Install arguments: /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1
    start /wait %PYTHON_INSTALLER_NAME% /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1
    
    if %errorlevel% neq 0 (
        echo Python installation failed. Please try installing Python manually.
        del %PYTHON_INSTALLER_NAME% > nul 2>&1
        pause
        exit /b 1
    )
    del %PYTHON_INSTALLER_NAME% > nul 2>&1
    echo Python installation complete. You might need to restart this script or command prompt for PATH changes to take full effect.
    echo Script will attempt to continue...
    
    :: Attempt to "refresh" PATH for the current session (limited effect)
    set PATH=%PATH% 
    
    :: Re-check Python (python.exe) as PATH might have been updated
    python --version > nul 2>&1
    if %errorlevel% neq 0 (
        echo Python (python.exe) still not found in current session after installation.
        echo Please close this window and re-run this script as Administrator.
        echo Alternatively, manually add Python install directory and its Scripts subdirectory to your system PATH.
        pause
        exit /b 1
    )
    set PYTHON_COMMAND=python
    goto python_env_checked

:python_env_checked
echo Python environment configured. Using: %PYTHON_COMMAND%
%PYTHON_COMMAND% --version
echo.

:: Check and install required Python libraries
echo Checking and installing necessary Python libraries (ensure %REQUIREMENTS_FILE% exists)...
if not exist "%REQUIREMENTS_FILE%" (
    echo ERROR: %REQUIREMENTS_FILE% not found. Cannot install dependencies.
    echo Please create %REQUIREMENTS_FILE% and list required libraries, for example:
    echo PySide6
    echo pandas
    echo requests
    pause
    exit /b 1
)

echo Installing/updating libraries using pip (this may take some time)...
%PYTHON_COMMAND% -m pip install --upgrade pip > nul
%PYTHON_COMMAND% -m pip install -r "%REQUIREMENTS_FILE%"
if %errorlevel% neq 0 (
    echo Failed to install/update libraries. Check pip and network connection.
    echo You can try running manually: %PYTHON_COMMAND% -m pip install -r "%REQUIREMENTS_FILE%"
    pause
    exit /b 1
)
echo Library check and installation complete.
echo.

:: Start the application
echo Starting application (%APP_FILE%)...
if not exist "%APP_FILE%" (
    echo ERROR: Application main file %APP_FILE% not found.
    pause
    exit /b 1
)

%PYTHON_COMMAND% "%APP_FILE%"

if %errorlevel% neq 0 (
    echo Application exited with an error. Error code: %errorlevel%
    pause
    exit /b 1
)

echo Application exited.
exit /b 0
