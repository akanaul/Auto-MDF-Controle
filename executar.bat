@echo off
setlocal enabledelayedexpansion

REM Simple robust launcher: prefer .venv python, else use py -3 or python on PATH
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "WRAPPER=%~dp0run_gerar_no_ui.py"
set "SCRIPT=%~dp0gerar_planilha.py"

echo [1/3] Solicitando nome do responsavel...
set /p RESPONSAVEL="Digite o nome do responsavel pela emissao: "
if "%RESPONSAVEL%"=="" (
    echo ERRO: Nome do responsavel e obrigatorio.
    pause
    exit /b 1
)
echo Responsavel definido: %RESPONSAVEL%

echo [2/3] Localizando interpretador Python...

REM Check .venv first
REM If .venv exists, use it. Otherwise attempt to create it and install deps.
if exist "%VENV_PY%" (
    echo [OK] Usando Python do .venv: %VENV_PY%
    if exist "%WRAPPER%" (
        "%VENV_PY%" "%WRAPPER%" "%RESPONSAVEL%"
    ) else (
        "%VENV_PY%" "%SCRIPT%" "%RESPONSAVEL%"
    )
    goto end
)

echo [INFO] .venv nao encontrado. Tentando criar ambiente virtual...

REM Try to create venv using py -3 or python if available
where py >nul 2>&1
if %errorlevel% equ 0 (
    echo [INFO] Criando .venv via py -3...
    py -3 -m venv "%~dp0.venv" >nul 2>&1
    if %errorlevel% equ 0 (
        echo [OK] .venv criado com py
    ) else (
        echo [WARN] falha ao criar .venv com py
    )
) else (
    where python >nul 2>&1
    if %errorlevel% equ 0 (
        echo [INFO] Criando .venv via python...
        python -m venv "%~dp0.venv" >nul 2>&1
        if %errorlevel% equ 0 (
            echo [OK] .venv criado com python
        ) else (
            echo [WARN] falha ao criar .venv com python
        )
    ) else (
        echo [WARN] Nao ha um interpretador Python disponivel para criar .venv
    )
)

REM If venv was created, try to install requirements.txt
if exist "%~dp0.venv\Scripts\pip.exe" (
    echo [INFO] Instalando dependencias do requirements.txt...
    "%~dp0.venv\Scripts\pip.exe" install -r "%~dp0requirements.txt" >nul 2>&1
    if %errorlevel% equ 0 (
        echo [OK] Dependencias instaladas no .venv
    ) else (
        echo [WARN] Falha ao instalar dependencias via pip (veja requirements.txt)
    )
) else (
    echo [INFO] pip do .venv nao encontrado; pulando instalacao de dependencias
)

REM If venv python exists now, prefer it and run the script inside the venv
if exist "%VENV_PY%" (
    echo [OK] Usando Python do .venv: %VENV_PY%
    if exist "%WRAPPER%" (
        "%VENV_PY%" "%WRAPPER%" "%RESPONSAVEL%"
    ) else (
        "%VENV_PY%" "%SCRIPT%" "%RESPONSAVEL%"
    )
    goto end
)

REM Check py launcher
where py >nul 2>&1
if %errorlevel% equ 0 (
    py -3 --version >nul 2>&1
    if !errorlevel! equ 0 (
        echo [OK] Usando py -3 launcher
        py -3 "%SCRIPT%" "%RESPONSAVEL%"
        goto end
    )
)

REM Check python in PATH
where python >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Usando python do PATH
    python "%SCRIPT%" "%RESPONSAVEL%"
    goto end
)

REM Search common Python installation directories
echo [INFO] Python nao encontrado no PATH. Procurando em diretorios comuns...

set "FOUND_PYTHON="
for %%V in (312 311 310 39 38) do (
    if exist "C:\Python%%V\python.exe" (
        set "FOUND_PYTHON=C:\Python%%V\python.exe"
        goto found_python
    )
)

REM Check AppData Local Programs
for %%V in (312 311 310 39 38) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set "FOUND_PYTHON=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        goto found_python
    )
)

REM Check Program Files
for %%V in (312 311 310 39 38) do (
    if exist "%ProgramFiles%\Python%%V\python.exe" (
        set "FOUND_PYTHON=%ProgramFiles%\Python%%V\python.exe"
        goto found_python
    )
    if exist "%ProgramFiles(x86)%\Python%%V\python.exe" (
        set "FOUND_PYTHON=%ProgramFiles(x86)%\Python%%V\python.exe"
        goto found_python
    )
)

REM Check Microsoft Store Python
if exist "%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe" (
    "%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe" --version >nul 2>&1
    if !errorlevel! equ 0 (
        set "FOUND_PYTHON=%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe"
        goto found_python
    )
)

REM Not found anywhere
echo ERRO: Nao foi possivel localizar um interpretador Python.
echo.
echo Opcoes:
echo  1. Instale Python de https://www.python.org/
echo  2. Execute o programa via VS Code
echo  3. Adicione Python ao PATH do sistema
goto end

:found_python
echo [OK] Python encontrado: !FOUND_PYTHON!
"!FOUND_PYTHON!" "%SCRIPT%" "%RESPONSAVEL%"

:end
pause
