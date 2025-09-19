@echo off

set SCRIPT="E:\Tasks\Payroll\Sick_Bank_Flag\src\main.py"
set LOGDIR="E:\Tasks\Payroll\Sick_Bank_Flag\log"
set LOG=%LOGDIR%\script_output.log

:: Ensure log directory exists
if not exist %LOGDIR% (
    mkdir %LOGDIR%
)

:: Start log with timestamp and user
echo ========================== > %LOG%
echo Start: %date% %time% >> %LOG%
whoami >> %LOG%

:: Run Python and capture output
echo Running Python script... >> %LOG%
"E:\Anaconda\python.exe" %SCRIPT% >> %LOG% 2>&1

:: End timestamp
echo End: %date% %time% >> %LOG%
echo ========================== >> %LOG%
