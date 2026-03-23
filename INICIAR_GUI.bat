@echo off
chcp 850 >nul
cd /d "%~dp0"

SET "PY_FULL="
SET "PYW_FULL="

:: -- 1. Le o caminho salvo pelo 1_INSTALAR.bat --------------------
IF EXIST "%~dp0python_path.txt" (
    SET /P PY_FULL=<"%~dp0python_path.txt"
)

:: Deriva o pythonw.exe do mesmo diretorio do python.exe salvo
IF NOT "%PY_FULL%"=="" (
    FOR %%D IN ("%PY_FULL%") DO SET "PY_DIR=%%~dpD"
    IF EXIST "%PY_DIR%pythonw.exe" SET "PYW_FULL=%PY_DIR%pythonw.exe"
)

:: -- 2. Fallback: busca dinamica (mesmo metodo do instalador) ------
IF "%PY_FULL%"=="" (
    py -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
    IF NOT ERRORLEVEL 1 (
        SET /P PY_FULL=<"%TEMP%\pypath.txt"
        DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
    )
)
IF "%PY_FULL%"=="" (
    python -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
    IF NOT ERRORLEVEL 1 (
        SET /P PY_FULL=<"%TEMP%\pypath.txt"
        DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
    )
)
IF "%PY_FULL%"=="" (
    FOR /D %%D IN ("%LOCALAPPDATA%\Programs\Python\*" "%LOCALAPPDATA%\Python\*") DO (
        IF "%PY_FULL%"=="" IF EXIST "%%D\python.exe" (
            "%%D\python.exe" -c "import sys; print(sys.executable)" >"%TEMP%\pypath.txt" 2>nul
            IF NOT ERRORLEVEL 1 (
                SET /P PY_FULL=<"%TEMP%\pypath.txt"
                DEL /Q "%TEMP%\pypath.txt" >nul 2>&1
            )
        )
    )
)

:: -- Detecta pythonw se ainda nao foi definido ---------------------
IF "%PYW_FULL%"=="" IF NOT "%PY_FULL%"=="" (
    FOR %%D IN ("%PY_FULL%") DO SET "PY_DIR=%%~dpD"
    IF EXIST "%PY_DIR%pythonw.exe" SET "PYW_FULL=%PY_DIR%pythonw.exe"
)

:: -- Inicia sem janela CMD -----------------------------------------
IF NOT "%PYW_FULL%"=="" (
    start "" "%PYW_FULL%" "%~dp0esocial_gui.py"
    exit /b 0
)
IF NOT "%PY_FULL%"=="" (
    start "" /B "%PY_FULL%" "%~dp0esocial_gui.py"
    exit /b 0
)

:: -- Nao encontrou -------------------------------------------------
echo.
echo ============================================================
echo   ERRO: Python nao encontrado
echo ============================================================
echo.
echo   Execute primeiro o arquivo: 1_INSTALAR.bat
echo.
pause
