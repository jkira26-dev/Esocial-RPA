@echo off
chcp 65001 >nul
TITLE DEV — eSocial RPA
cd /d "%~dp0"

echo ============================================================
echo   eSocial RPA — MODO DESENVOLVEDOR
echo ============================================================
echo.
echo  Ambiente detectado:
py --version 2>&1
echo.
echo  Dependencias:
py -m pip show playwright plyer 2>&1 | findstr /C:"Name:" /C:"Version:"
echo.
echo ============================================================
echo   Escolha o que executar:
echo   [1] Abrir GUI  (esocial_gui.py)
echo   [2] Abrir RPA CLI  (esocial_rpa.py)
echo   [3] Verificar dependencias completo
echo   [4] Checar sintaxe dos arquivos
echo   [5] Abrir Chrome com debug (porta 9222)
echo   [0] Sair
echo ============================================================
echo.
set /p OPCAO="  Opcao: "

if "%OPCAO%"=="1" goto :GUI
if "%OPCAO%"=="2" goto :CLI
if "%OPCAO%"=="3" goto :DEPS
if "%OPCAO%"=="4" goto :SYNTAX
if "%OPCAO%"=="5" goto :CHROME
if "%OPCAO%"=="0" exit /b 0
echo Opcao invalida.
pause
goto :EOF

:GUI
echo.
echo  Iniciando GUI (com console visivel para ver erros)...
echo  Feche a janela da GUI para retornar aqui.
echo.
py esocial_gui.py
echo.
echo  GUI encerrada. Pressione qualquer tecla para voltar ao menu.
pause >nul
goto :EOF

:CLI
echo.
echo  Iniciando RPA CLI...
py esocial_rpa.py
pause
goto :EOF

:DEPS
echo.
echo === Dependencias instaladas ===
py -m pip show playwright plyer greenlet pyee playwright-stealth 2>&1
echo.
echo === Browser Playwright ===
py -c "from playwright.sync_api import sync_playwright; p=sync_playwright().start(); b=p.chromium.launch(headless=True); print('  Chromium: OK  versao=' + b.version); b.close(); p.stop()" 2>&1
echo.
pause
goto :EOF

:SYNTAX
echo.
echo === Verificando sintaxe ===
py -c "import py_compile; [print(f'  {f}: OK') if not py_compile.compile(f, doraise=True) else None for f in ['config.py','esocial_rpa.py','esocial_gui.py']]" 2>&1
echo.
pause
goto :EOF

:CHROME
echo.
echo  Abrindo Google Chrome com debugger remoto na porta 9222...
echo  Faca login no eSocial e entao execute o robo.
echo.
set "CHROME_EXE="
for %%C in (
    "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
) do (
    if "!CHROME_EXE!"=="" if exist %%C set "CHROME_EXE=%%C"
)
setlocal enabledelayedexpansion
for %%C in (
    "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
) do (
    if "!CHROME_EXE!"=="" if exist %%C set "CHROME_EXE=%%C"
)
if "!CHROME_EXE!"=="" (
    echo  ERRO: Chrome nao encontrado. Instale o Google Chrome.
) else (
    echo  Abrindo: !CHROME_EXE!
    start "" "!CHROME_EXE!" --remote-debugging-port=9222 --user-data-dir="%TEMP%\esocial_chrome_dev"
    echo  Chrome aberto. Aguarde carregar e faca o login no eSocial.
)
endlocal
echo.
pause
goto :EOF
