@echo off
chcp 850 >nul
TITLE eSocial RPA - Desinstalador
color 0C
cd /d "%~dp0"

echo.
echo ============================================================
echo   eSocial RPA - Desinstalador
echo ============================================================
echo.
echo ATENCAO: O programa sera removido do computador.
echo          Os arquivos baixados em C:\esocial_xmls serao mantidos.
echo.
SET /P "CONFIRMA=Deseja continuar com a desinstalacao? (S/N): "
IF /I NOT "%CONFIRMA%"=="S" (
    echo Desinstalacao cancelada.
    pause
    exit /b 0
)

net session >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Necessita permissao de Administrador.
    pause
    exit /b 1
)

SET "INSTALL_DIR=%ProgramFiles%\eSocial RPA"
SET "STARTMENU=%ProgramData%\Microsoft\Windows\Start Menu\Programs\eSocial RPA"
SET "UNINSTALL_KEY=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\eSocialRPA"
FOR /F "tokens=2*" %%A IN ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v Desktop 2^>nul') DO SET "USER_DESKTOP=%%B"

echo.
echo [1/4] Removendo atalhos...
IF EXIST "%PUBLIC%\Desktop\eSocial RPA.lnk"  DEL /Q "%PUBLIC%\Desktop\eSocial RPA.lnk"
IF EXIST "%PUBLIC%\Desktop\eSocial RPA.url"  DEL /Q "%PUBLIC%\Desktop\eSocial RPA.url"
IF EXIST "%USER_DESKTOP%\eSocial RPA.lnk"    DEL /Q "%USER_DESKTOP%\eSocial RPA.lnk"
IF EXIST "%USER_DESKTOP%\eSocial RPA.url"    DEL /Q "%USER_DESKTOP%\eSocial RPA.url"
IF EXIST "%STARTMENU%"                       RD /S /Q "%STARTMENU%"

echo [2/4] Removendo registro...
reg delete "%UNINSTALL_KEY%" /f >nul 2>&1

echo [3/4] Agendando remocao dos arquivos...
SET "DEL_SCRIPT=%TEMP%\esocial_del.bat"
(
    echo @echo off
    echo ping 127.0.0.1 -n 3 ^>nul
    echo IF EXIST "%INSTALL_DIR%" RD /S /Q "%INSTALL_DIR%"
    echo DEL /Q "%~f0"
) > "%DEL_SCRIPT%"

echo [4/4] Concluindo...
echo.
echo ============================================================
echo   Desinstalacao concluida.
echo ============================================================
echo.
echo   Mantidos (seus dados):
echo     - C:\esocial_xmls\           (XMLs baixados)
echo     - %APPDATA%\eSocialRPA\      (configuracoes e progresso)
echo.

start /B cmd /C "%DEL_SCRIPT%"
pause
exit /b 0
