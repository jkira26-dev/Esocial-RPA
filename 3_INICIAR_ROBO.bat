@echo off
TITLE eSocial RPA - Robo de Download XML

echo ============================================================
echo   eSocial RPA - Iniciando Robo
echo ============================================================
echo.
echo   ATENCAO: O Chrome precisa estar aberto com o login
echo   do eSocial ja realizado. Se ainda nao fez isso,
echo   feche esta janela e execute primeiro: 2_ABRIR_CHROME.bat
echo.

cd /d "%~dp0"

python esocial_rpa.py

echo.
echo ============================================================
echo   Robo finalizado. Verifique o arquivo esocial_rpa.log
echo   para detalhes da execucao.
echo ============================================================
pause
