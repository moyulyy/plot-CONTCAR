@echo off
rem 打包便携版单文件 exe（需要先在 chem_env 安装 pyinstaller）
chcp 65001 >nul
cd /d "%~dp0"
set "PY=D:\miniconda3\envs\chem_env\python.exe"
rem conda 环境下 pyexpat 依赖的 libexpat.dll 需要手动打入
set "EXPAT=D:\miniconda3\envs\chem_env\Library\bin\libexpat.dll"
if not exist "%PY%" set "PY=python"
"%PY%" -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name plot-CONTCAR ^
  --icon assets\app.ico ^
  --add-data "assets\app.ico;assets" ^
  --add-data "assets\app.png;assets" ^
  --add-binary "%EXPAT%;." ^
  gui_launcher.pyw
echo.
echo ============================================
echo  已生成： dist\plot-CONTCAR.exe
echo ============================================
pause
