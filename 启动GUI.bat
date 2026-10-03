@echo off
rem 无控制台启动（优先 chem_env 的 pythonw，找不到则用 PATH 里的 pythonw）
cd /d "%~dp0"
set "PYW=D:\miniconda3\envs\chem_env\pythonw.exe"
if not exist "%PYW%" set "PYW=pythonw"
start "" "%PYW%" "%~dp0gui_launcher.pyw"
exit
