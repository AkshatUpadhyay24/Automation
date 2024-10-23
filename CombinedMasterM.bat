@echo off
for /f "delims=" %%i in ('where python') do set python_path=%%i
%python_path% "D:\Automation\Scripts\CombinedMasterM.py"
pause