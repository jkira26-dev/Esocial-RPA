@echo off
TITLE eSocial RPA - Instalador Unattended
cd /d "%~dp0"

echo [1/5] Criando pastas (Modo Unattended em HKCU/AppData)...
SET "INSTALL_DIR=%LOCALAPPDATA%\eSocial RPA Test"
SET "DATA_DIR=%LOCALAPPDATA%\eSocialRPA TestData"
SET "STARTMENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs\eSocial RPA Test"
SET "UNINSTALL_KEY=HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\eSocialRPA"

IF NOT EXIST "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
IF NOT EXIST "%DATA_DIR%"    mkdir "%DATA_DIR%"
IF NOT EXIST "%STARTMENU%"   mkdir "%STARTMENU%"

echo [2/5] Copiando arquivos...
FOR %%F IN ("%~dp0*") DO (
    IF /I NOT "%%~nxF"=="INSTALAR.bat" (
        IF /I NOT "%%~nxF"=="DESINSTALAR.bat" (
            copy /Y "%%F" "%INSTALL_DIR%\" >nul 2>&1
        )
    )
)
IF EXIST "%~dp0dist\" xcopy /E /I /Y /Q "%~dp0dist\" "%INSTALL_DIR%\dist\" >nul 2>&1
IF NOT EXIST "%DATA_DIR%\config.py" (
    IF EXIST "%~dp0config.py" copy /Y "%~dp0config.py" "%DATA_DIR%\config.py" >nul 2>&1
)
copy /Y "%~dp0DESINSTALAR.bat" "%INSTALL_DIR%\DESINSTALAR.bat" >nul 2>&1

echo [3/5] Criando atalhos...
SET "PY_FULL="
IF EXIST "%~dp0python_path.txt" SET /P PY_FULL=<"%~dp0python_path.txt"
IF "%PY_FULL%"=="" (
    py -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
    IF NOT ERRORLEVEL 1 SET /P PY_FULL=<"%TEMP%\pypath.txt"
    DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
)
IF "%PY_FULL%"=="" SET "PY_FULL=python"

FOR %%D IN ("%PY_FULL%") DO SET "PY_DIR=%%~dpD"
SET "PYW_FULL=%PY_DIR%pythonw.exe"
IF NOT EXIST "%PYW_FULL%" SET "PYW_FULL=%PY_FULL%"

SET "PS_SCRIPT=%TEMP%\esocial_sc.ps1"
(
echo $sh = New-Object -ComObject WScript.Shell
echo $userDesktop = ^(Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'Desktop'^).Desktop
echo function CriarAtalho($path^) {
echo   $s = $sh.CreateShortcut($path^)
echo   $s.TargetPath = '%PYW_FULL%'
echo   $s.Arguments = '"%INSTALL_DIR%\esocial_gui.py"'
echo   $s.WorkingDirectory = '%INSTALL_DIR%'
echo   $s.Description = 'eSocial RPA'
echo   if (Test-Path '%INSTALL_DIR%\icon.ico'^) { $s.IconLocation = '%INSTALL_DIR%\icon.ico' }
echo   $s.Save(^)
echo }
echo CriarAtalho("$userDesktop\eSocial RPA Test.lnk"^)
echo CriarAtalho('%STARTMENU%\eSocial RPA Test.lnk'^)
echo $u = $sh.CreateShortcut('%STARTMENU%\Desinstalar eSocial RPA Test.lnk'^)
echo $u.TargetPath = '%INSTALL_DIR%\DESINSTALAR.bat'
echo $u.WorkingDirectory = '%INSTALL_DIR%'
echo $u.Save(^)
) > "%PS_SCRIPT%"

REM powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

echo [4/5] Registrando no sistema...
reg add "%UNINSTALL_KEY%" /v "DisplayName"     /t REG_SZ /d "eSocial RPA" /f >nul 2>&1
reg add "%UNINSTALL_KEY%" /v "DisplayVersion"  /t REG_SZ /d "1.0.0" /f >nul 2>&1
reg add "%UNINSTALL_KEY%" /v "Publisher"       /t REG_SZ /d "IOB Gestao Contabil" /f >nul 2>&1
reg add "%UNINSTALL_KEY%" /v "InstallLocation" /t REG_SZ /d "%INSTALL_DIR%" /f >nul 2>&1
reg add "%UNINSTALL_KEY%" /v "UninstallString" /t REG_SZ /d "%INSTALL_DIR%\DESINSTALAR.bat" /f >nul 2>&1
reg add "%UNINSTALL_KEY%" /v "NoModify"        /t REG_DWORD /d 1 /f >nul 2>&1
reg add "%UNINSTALL_KEY%" /v "NoRepair"        /t REG_DWORD /d 1 /f >nul 2>&1
IF EXIST "%INSTALL_DIR%\icon.ico" reg add "%UNINSTALL_KEY%" /v "DisplayIcon" /t REG_SZ /d "%INSTALL_DIR%\icon.ico" /f >nul 2>&1

echo [5/5] Instalando dependencias Python...
IF NOT "%PY_FULL%"=="python" (
    "%PY_FULL%" -m pip install playwright plyer --no-warn-script-location --quiet >nul 2>&1
    "%PY_FULL%" -m playwright install chromium >nul 2>&1
    echo       OK.
)

echo.
echo ============================================================
echo   Instalacao Unattended Concluida!
echo ============================================================
exit /b 0
