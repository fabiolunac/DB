@echo off
echo ===========================================
echo   Gerando executavel com PyInstaller
echo ===========================================

REM Ativar o venv
call venv\Scripts\activate.bat

REM Ajuste aqui o nome do seu script Python
set SCRIPT=add_line.py

REM Nome do executavel final
set NAME=add_line

REM Executa o PyInstaller
pyinstaller --onefile --name %NAME% %SCRIPT%

echo.
echo ===========================================
echo   Concluido! O executavel esta em /dist
echo ===========================================
pause