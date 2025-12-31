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
if exist "%VENV_PY%" (
    echo [OK] Usando Python do .venv: %VENV_PY%
    if exist "%WRAPPER%" (
        "%VENV_PY%" "%WRAPPER%" "%RESPONSAVEL%"
    ) else (
        "%VENV_PY%" "%SCRIPT%" "%RESPONSAVEL%"
    )
    goto end
)

py -3 "%SCRIPT%" "%RESPONSAVEL%" >nul 2>&1 && (
    echo [OK] Usando py -3 launcher
    py -3 "%SCRIPT%" "%RESPONSAVEL%"
    goto end
)

python "%SCRIPT%" "%RESPONSAVEL%" >nul 2>&1 && (
    echo [OK] Usando python do PATH
    python "%SCRIPT%" "%RESPONSAVEL%"
    goto end
)

echo ERRO: Nao foi possivel localizar um interpretador Python. Instale o Python e tente novamente.
echo Recomendo executar o programa via VS Code ou instalar Python de https://www.python.org/

:end
pause
