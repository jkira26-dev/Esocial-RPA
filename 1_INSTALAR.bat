@echo off
chcp 850 >nul
TITLE Instalacao - eSocial RPA
cd /d "%~dp0"

echo ============================================================
echo   eSocial RPA -- Instalacao de Dependencias
echo ============================================================
echo.
echo Buscando Python instalado...

SET "PY_FULL="

:: -- Estrategia: obtem o caminho COMPLETO do executavel Python ----
:: Isso garante que usamos o mesmo python.exe para pip, playwright etc.

:: 1. Tenta resolver pelo launcher "py" (mais confiavel no Windows)
py -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
IF NOT ERRORLEVEL 1 (
    SET /P PY_FULL=<"%TEMP%\pypath.txt"
    DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
    GOTO :ENCONTROU
)

:: 2. Tenta "python" no PATH (ignorando alias da Store via teste real)
python -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
IF NOT ERRORLEVEL 1 (
    SET /P PY_FULL=<"%TEMP%\pypath.txt"
    DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
    GOTO :ENCONTROU
)

:: 3. Busca manual em pastas comuns
FOR %%C IN (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "%ProgramFiles%\Python313\python.exe"
    "%ProgramFiles%\Python312\python.exe"
    "%ProgramFiles%\Python311\python.exe"
) DO (
    IF "%PY_FULL%"=="" IF EXIST %%C (
        %%C -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
        IF NOT ERRORLEVEL 1 (
            SET /P PY_FULL=<"%TEMP%\pypath.txt"
            DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
        )
    )
)

:: 4. Busca por glob em AppData (cobre pythoncore e outras variantes)
FOR /D %%D IN ("%LOCALAPPDATA%\Programs\Python\*") DO (
    IF "%PY_FULL%"=="" IF EXIST "%%D\python.exe" (
        "%%D\python.exe" -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
        IF NOT ERRORLEVEL 1 (
            SET /P PY_FULL=<"%TEMP%\pypath.txt"
            DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
        )
    )
)
FOR /D %%D IN ("%LOCALAPPDATA%\Python\*") DO (
    IF "%PY_FULL%"=="" IF EXIST "%%D\python.exe" (
        "%%D\python.exe" -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
        IF NOT ERRORLEVEL 1 (
            SET /P PY_FULL=<"%TEMP%\pypath.txt"
            DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
        )
    )
)

IF "%PY_FULL%"=="" GOTO :NAO_ENCONTROU
GOTO :ENCONTROU

:ENCONTROU
echo.
echo [OK] Python encontrado: %PY_FULL%
"%PY_FULL%" --version
echo.

echo [1/3] Atualizando pip...
"%PY_FULL%" -m pip install --upgrade pip --no-warn-script-location --quiet
IF ERRORLEVEL 1 (
    echo [AVISO] Nao foi possivel atualizar o pip. Continuando...
)

echo [2/3] Instalando playwright e plyer...
"%PY_FULL%" -m pip install playwright plyer --no-warn-script-location
IF ERRORLEVEL 1 (
    echo.
    echo [ERRO] Falha ao instalar dependencias.
    echo        Possiveis causas:
    echo        - Sem conexao com a internet
    echo        - Permissao negada ^(tente como Administrador^)
    echo.
    pause
    exit /b 1
)

echo.
echo [3/3] Instalando navegador Chromium...
"%PY_FULL%" -m playwright install chromium
IF ERRORLEVEL 1 (
    echo [AVISO] Chromium nao instalado. O robo usara o Chrome do Windows.
)

:: Salva o caminho do Python para uso pelo INICIAR_GUI.bat
echo %PY_FULL%> "%~dp0python_path.txt"
echo [INFO] Caminho do Python salvo em python_path.txt

echo.
echo ============================================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo ============================================================
echo.
echo   Execute INICIAR_GUI.bat para abrir o programa.
echo.
pause
exit /b 0

:NAO_ENCONTROU
echo.
echo ============================================================
echo   [ERRO] Python nao encontrado
echo ============================================================
echo.
echo   SOLUCAO 1 - Instalar o Python oficial ^(recomendado^):
echo     1. Acesse https://www.python.org/downloads/
echo     2. Baixe Python 3.11 ou superior
echo     3. Durante a instalacao marque "Add Python to PATH"
echo     4. Reinicie o computador e execute este arquivo novamente
echo.
echo   SOLUCAO 2 - Desabilitar alias da Microsoft Store:
echo     Configuracoes ^> Aplicativos ^> Configuracoes avancadas
echo     Aliases de execucao do aplicativo
echo     Desative: python.exe e python3.exe
echo.
pause
exit /b 1
