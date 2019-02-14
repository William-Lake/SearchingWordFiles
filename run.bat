@ECHO off

REM  ======================================================================= Main

REM We first need to check that Python is installed.
python -V

REM If True, Python is not installed.
IF ERRORLEVEL 1 GOTO NoPython

REM Then we need to ensure the dependencies are installed.

FOR /F "tokens=*" %%L IN (requirements.txt) DO CALL :CheckForModule %%L

REM Then we can start the program
python search_word_files

GOTO :EOF

REM  ======================================================================= NoPython

REM Outlines what to do if Python isn't installed

:NoPython
ECHO Python is not installed and needs to be both installed and on your PATH in order to use this application.
set /p tmp=Press any key to open a web browser to the Python Downloads section.
REM start "" https://www.python.org/downloads/
ECHO.
ECHO Be sure to select the "Add Python to PATH" checkbox at the bottom of the first page.
ECHO When you've finished, re-run this script to continue.
ECHO.
ECHO For more info about running Python on Windows, see here: https://docs.python.org/3/using/windows.html

GOTO :EOF

REM  ======================================================================= CheckForModule

REM Ensures a given Python Module is installed.

:CheckForModule
IF "%~1%" == "pywin32" (
    python -c "import win32com.client as win32"
) ELSE (
    python -c "import %~1%"
)

IF ERRORLEVEL 1 CALL :NoModule %~1%

GOTO :EOF

REM  ======================================================================= NoModule

REM Attempts to install module, asking user first.

:NoModule
SET /p do_install = The Python module %~1% is not installed and is required by this application. Install it? [Y/N]
IF /I "%do_install%" NEQ "Y"  GOTO :EOF
pip install %~1%

GOTO :EOF
