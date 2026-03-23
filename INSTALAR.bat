@echo off
chcp 850 >nul
TITLE eSocial RPA - Instalador
color 0A
cd /d "%~dp0"

echo.
echo ============================================================
echo   eSocial RPA - Instalador
echo   Download Automatico de XMLs do eSocial
echo ============================================================
echo.
echo Este instalador ira:
echo   1. Copiar os arquivos para Arquivos de Programas
echo   2. Criar atalho na area de trabalho
echo   3. Criar atalho no Menu Iniciar
echo   4. Registrar o desinstalador no Windows
echo.
echo Pressione qualquer tecla para continuar ou feche para cancelar.
pause >nul

:: Verifica Administrador
net session >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [AVISO] Este instalador precisa de permissao de Administrador.
    echo         Clique com botao direito e "Executar como administrador".
    echo.
    pause
    exit /b 1
)

:: Caminhos
SET "INSTALL_DIR=%ProgramFiles%\eSocial RPA"
SET "DATA_DIR=%APPDATA%\eSocialRPA"
SET "STARTMENU=%ProgramData%\Microsoft\Windows\Start Menu\Programs\eSocial RPA"
SET "UNINSTALL_KEY=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\eSocialRPA"

FOR /F "tokens=2*" %%A IN ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v Desktop 2^>nul') DO SET "USER_DESKTOP=%%B"

echo.
echo [1/5] Criando pastas...
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

:: Descobre Python real (mesmo metodo do 1_INSTALAR.bat)
SET "PY_FULL="
IF EXIST "%~dp0python_path.txt" SET /P PY_FULL=<"%~dp0python_path.txt"
IF "%PY_FULL%"=="" (
    py -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
    IF NOT ERRORLEVEL 1 SET /P PY_FULL=<"%TEMP%\pypath.txt"
    DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
)
IF "%PY_FULL%"=="" (
    python -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
    IF NOT ERRORLEVEL 1 SET /P PY_FULL=<"%TEMP%\pypath.txt"
    DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
)
IF "%PY_FULL%"=="" SET "PY_FULL=python"

:: pythonw no mesmo diretorio
FOR %%D IN ("%PY_FULL%") DO SET "PY_DIR=%%~dpD"
SET "PYW_FULL=%PY_DIR%pythonw.exe"
IF NOT EXIST "%PYW_FULL%" SET "PYW_FULL=%PY_FULL%"

SET "PS_SCRIPT=%TEMP%\esocial_sc.ps1"
(
echo $sh = New-Object -ComObject WScript.Shell
echo function CriarAtalho($path^) {
echo   $s = $sh.CreateShortcut($path^)
echo   $s.TargetPath = '%PYW_FULL%'
echo   $s.Arguments = '"%INSTALL_DIR%\esocial_gui.py"'
echo   $s.WorkingDirectory = '%INSTALL_DIR%'
echo   $s.Description = 'eSocial RPA'
echo   if (Test-Path '%INSTALL_DIR%\icon.ico'^) { $s.IconLocation = '%INSTALL_DIR%\icon.ico' }
echo   $s.Save(^)
echo }
echo CriarAtalho('%PUBLIC%\Desktop\eSocial RPA.lnk'^)
echo CriarAtalho('%USER_DESKTOP%\eSocial RPA.lnk'^)
echo CriarAtalho('%STARTMENU%\eSocial RPA.lnk'^)
echo $u = $sh.CreateShortcut('%STARTMENU%\Desinstalar eSocial RPA.lnk'^)
echo $u.TargetPath = '%INSTALL_DIR%\DESINSTALAR.bat'
echo $u.WorkingDirectory = '%INSTALL_DIR%'
echo $u.Save(^)
) > "%PS_SCRIPT%"
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" >nul 2>&1
DEL /Q "%PS_SCRIPT%" >nul 2>&1

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
) ELSE (
    echo       Execute 1_INSTALAR.bat para instalar as dependencias.
)

IF NOT EXIST "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    IF NOT EXIST "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
        echo.
        echo [ATENCAO] Google Chrome nao encontrado.
        echo           Instale em: https://www.google.com/chrome
    )
)

echo.
echo ============================================================
echo   Instalacao concluida!
echo ============================================================
echo.
echo   Atalhos criados em:
echo     - Area de trabalho
echo     - Menu Iniciar ^> eSocial RPA
echo.
echo   PROXIMO PASSO:
echo     1. Abra o eSocial RPA pela area de trabalho
echo     2. Clique em "Abrir Chrome" e faca o login no eSocial
echo     3. Clique em "Verificar" para confirmar a conexao
echo.
SET /P "ABRIR=Deseja abrir o eSocial RPA agora? (S/N): "
IF /I "%ABRIR%"=="S" start "" "%INSTALL_DIR%\INICIAR_GUI.bat"
echo.
pause
exit /b 0
