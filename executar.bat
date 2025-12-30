@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo.
echo ============================================================
echo GERADOR DE PLANILHA DE MDFs - LAUNCHER
echo ============================================================
echo.

REM Define o caminho do projeto (com aspas para caminhos com espacos)
set "PROJECT_DIR=%~dp0"
set "VENV_DIR=%PROJECT_DIR%.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "PIP_EXE=%VENV_DIR%\Scripts\pip.exe"
set "REQUIREMENTS_FILE=%PROJECT_DIR%requirements.txt"

REM ============================================================
REM Funcao para encontrar Python
REM ============================================================
set PYTHON_CMD=
set PYTHON_FOUND=0

REM Tenta usar py.exe (Python Launcher para Windows)
py -3 --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=py -3"
    set PYTHON_FOUND=1
    echo [OK] Python encontrado via py.exe
    goto python_found
)

REM Tenta python no PATH
python --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python"
    set PYTHON_FOUND=1
    echo [OK] Python encontrado no PATH
    goto python_found
)

REM Tenta python3 no PATH
python3 --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python3"
    set PYTHON_FOUND=1
    echo [OK] Python encontrado (python3) no PATH
    goto python_found
)

REM Se nao encontrou nada
if !PYTHON_FOUND! equ 0 (
    echo.
    echo ============================================================
    echo ERRO: Python nao foi encontrado no sistema!
    echo ============================================================
    echo.
    echo Este computador nao possui Python instalado ou configurado
    echo corretamente no PATH do sistema.
    echo.
    echo SOLUCAO:
    echo 1. Baixe Python em: https://www.python.org/downloads/
    echo 2. Execute o instalador
    echo 3. IMPORTANTE: Marque a opcao "Add Python to PATH" durante instalacao
    echo 4. Reinicie o computador
    echo 5. Execute novamente este arquivo
    echo.
    echo Ou, se ja tem Python instalado mas nao no PATH:
    echo 1. Localize onde Python esta instalado
    echo 2. Adicione o diretorio ao PATH do sistema (Propriedades do Sistema)
    echo.
    pause
    exit /b 1
)

:python_found
echo.

REM Verifica se o venv existe
if not exist "%VENV_DIR%" (
    echo [1/3] Criando ambiente virtual...
    %PYTHON_CMD% -m venv "%VENV_DIR%"
    if %ERRORLEVEL% NEQ 0 (
        echo ERRO: Nao foi possivel criar o ambiente virtual.
        pause
        exit /b 1
    )
    echo [OK] Ambiente virtual criado
    echo.
)

REM Depois de criar/garantir venv, atualizar caminhos do executavel
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "PIP_EXE=%VENV_DIR%\Scripts\pip.exe"

REM Verifica e instala dependências
echo [2/3] Verificando dependencias...

REM Verifica se requirements.txt existe
if exist "%REQUIREMENTS_FILE%" (
    echo Instalando pacotes de %REQUIREMENTS_FILE%...
    if exist "%PIP_EXE%" (
        "%PIP_EXE%" install -q -r "%REQUIREMENTS_FILE%"
    ) else (
        "%PYTHON_EXE%" -m pip install -q -r "%REQUIREMENTS_FILE%"
    )
    if %ERRORLEVEL% NEQ 0 (
        echo AVISO: Alguns pacotes podem nao ter sido instalados corretamente.
    )
) else (
    echo Instalando dependencias padrao...
    if exist "%PIP_EXE%" (
        "%PIP_EXE%" install -q pandas pdfplumber openpyxl
    ) else (
        "%PYTHON_EXE%" -m pip install -q pandas pdfplumber openpyxl
    )
    if %ERRORLEVEL% NEQ 0 (
        echo AVISO: Alguns pacotes podem nao ter sido instalados corretamente.
    )
)
echo [OK] Dependencias verificadas
echo.

REM Executa o programa
echo [3/3] Iniciando programa...
echo.
"%PYTHON_EXE%" "%PROJECT_DIR%gerar_planilha.py" %*

pause
