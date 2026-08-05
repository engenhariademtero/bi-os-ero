@echo off
call "%~dp0sincronizar.bat"
echo.
if errorlevel 1 (echo A atualizacao falhou. Consulte sincronizacao.log.) else (echo Atualizacao concluida.)
pause
