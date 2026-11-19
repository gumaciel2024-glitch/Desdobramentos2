@echo off
setlocal
title Triagem SISBAJUD
cd /d "%~dp0" || goto :erro
if not exist "triagem.py" (
    echo Extraia todos os arquivos do ZIP antes de abrir iniciar.bat.
    goto :erro
)
if not exist ".venv\Scripts\python.exe" (
    echo Preparando o ambiente Python...
    py -3 -m venv ".venv"
    if errorlevel 1 (
        echo Verifique se o Python 3.11 ou superior esta instalado com o Python Launcher.
        goto :erro
    )
)
set "ROBO_PY=.venv\Scripts\python.exe"
if not exist ".venv\sisbajud-pronto" (
    echo Instalando as dependencias do robo...
    "%ROBO_PY%" -m pip install "playwright>=1.51" openpyxl
    if errorlevel 1 goto :erro
    "%ROBO_PY%" -m playwright install chromium
    if errorlevel 1 goto :erro
    type nul > ".venv\sisbajud-pronto"
)
"%ROBO_PY%" "triagem.py"
if errorlevel 1 goto :erro
pause
exit /b 0

:erro
echo.
echo A abertura ou a execucao falhou. Veja a mensagem acima.
pause
exit /b 1
