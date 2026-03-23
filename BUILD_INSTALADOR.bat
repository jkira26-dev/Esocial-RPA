@echo off
TITLE eSocial RPA — Build do Instalador
color 0A
cd /d "%~dp0"

echo.
echo ====================================================
echo   eSocial RPA — Gerando instalador Windows
echo ====================================================
echo.

:: ── Passo 1: Verifica Python ────────────────────────
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Python nao encontrado. Instale em https://python.org
    pause & exit /b 1
)

:: ── Passo 2: Instala PyInstaller se necessário ──────
echo [1/4] Verificando PyInstaller...
python -m PyInstaller --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo       Instalando PyInstaller...
    python -m pip install pyinstaller --quiet
)
echo       OK.

:: ── Passo 3: Instala dependências ───────────────────
echo [2/4] Instalando dependencias...
python -m pip install playwright plyer --quiet
python -m playwright install chromium --quiet
echo       OK.

:: ── Passo 4: Gera executável com PyInstaller ────────
echo [3/4] Gerando executavel (PyInstaller)...
python -m PyInstaller esocial_rpa.spec --noconfirm --clean
IF %ERRORLEVEL% NEQ 0 (
    echo [ERRO] PyInstaller falhou. Verifique os logs acima.
    pause & exit /b 1
)
echo       Executavel gerado em: dist\eSocialRPA\

:: ── Passo 5: Compila instalador com Inno Setup ──────
echo [4/4] Compilando instalador (Inno Setup)...

SET INNO1=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
SET INNO2=C:\Program Files\Inno Setup 6\ISCC.exe

IF EXIST "%INNO1%" (
    "%INNO1%" installer.iss
    GOTO inno_ok
)
IF EXIST "%INNO2%" (
    "%INNO2%" installer.iss
    GOTO inno_ok
)

echo.
echo [AVISO] Inno Setup nao encontrado.
echo         Baixe em: https://jrsoftware.org/isdl.php
echo         Apos instalar, execute este script novamente.
echo.
echo         O executavel foi gerado em: dist\eSocialRPA\
echo         Voce pode distribuir essa pasta diretamente.
echo.
goto fim

:inno_ok
echo       OK.
echo.
echo ====================================================
echo   Instalador gerado em: installer_output\
echo ====================================================

:fim
echo.
pause
