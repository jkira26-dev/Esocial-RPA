@echo off
title BUILD eSocial RPA
echo ========================================================
echo   BUILD eSocial RPA — NOVAL ARQUITETURA
echo ========================================================
echo.

:: Verifica Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado no PATH.
    pause
    exit /b
)

:: Instala dependencias necessarias
echo [1/4] Instalando dependencias...
pip install playwright pywin32 plyer win10toast pyinstaller --quiet

:: Instala o browser do Playwright
echo [2/4] Verificando browser do Playwright...
playwright install chromium

:: Executa o Build
echo [3/4] Gerando executavel (Dist/)...
pyinstaller esocial_rpa.spec --noconfirm

echo.
echo [4/4] Concluido!
echo O executavel esta na pasta: dist\eSocialRPA
echo.
pause
