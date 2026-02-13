@echo off
echo ========================================================
echo   Limpando versoes antigas e preparando compilacao...
echo ========================================================

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /f /q *.spec

echo.
echo ========================================================
echo   Iniciando Compilacao com ICONE...
echo ========================================================

:: O parametro --icon adiciona a imagem ao .exe
pyinstaller --clean --noconsole --onefile --icon=imagemvauto.ico --collect-all customtkinter --collect-all matplotlib VAuto.py

echo.
echo ========================================================
echo   PROCESSO CONCLUIDO! Verifique a pasta DIST.
echo ========================================================
pause