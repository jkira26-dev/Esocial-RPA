@echo off
TITLE Abrindo Chrome com Debugger - eSocial

echo ============================================================
echo   PASSO 1: Abrindo Chrome com porta de depuracao
echo ============================================================
echo.
echo   Isso e necessario para que o robo possa se conectar
echo   ao Chrome onde voce fara o login com certificado.
echo.

REM Tenta localizar o Chrome em caminhos comuns do Windows
SET CHROME_PATH=""

IF EXIST "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    SET CHROME_PATH="C:\Program Files\Google\Chrome\Application\chrome.exe"
) ELSE IF EXIST "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    SET CHROME_PATH="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
) ELSE IF EXIST "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
    SET CHROME_PATH="%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
)

IF %CHROME_PATH%=="" (
    echo [ERRO] Chrome nao encontrado nos caminhos padrao.
    echo Edite este arquivo e coloque o caminho correto do chrome.exe
    pause
    exit /b 1
)

REM Cria pasta de perfil separada para nao interferir no Chrome do dia a dia
IF NOT EXIST "C:\chrome_esocial" mkdir "C:\chrome_esocial"

echo   Abrindo Chrome...
echo   (Uma janela nova do Chrome sera aberta)
echo.

start "" %CHROME_PATH% --remote-debugging-port=9222 --user-data-dir="C:\chrome_esocial" --no-first-run --no-default-browser-check https://esocial.gov.br

echo ============================================================
echo   Chrome aberto! 
echo ============================================================
echo.
echo   AGORA VOCE DEVE:
echo   1. No Chrome que abriu, faca o login no eSocial
echo      com o certificado digital normalmente
echo.
echo   2. Selecione o perfil (Empregador ou Procurador)
echo.
echo   3. Quando estiver na tela inicial do eSocial,
echo      volte aqui e execute o arquivo:
echo      2_INICIAR_ROBO.bat
echo.
pause
