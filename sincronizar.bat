@echo off
cd /d "%~dp0"
py -3 atualizar_dados.py
if errorlevel 1 python atualizar_dados.py
exit /b %errorlevel%
