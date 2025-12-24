# Script de Launcher para Gerador de Planilha de MDFs
# Gerencia ambiente virtual, dependências e execução do programa

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

# Cores para output
$colors = @{
    'header' = 'Cyan'
    'success' = 'Green'
    'warning' = 'Yellow'
    'error' = 'Red'
    'info' = 'White'
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Type = 'info'
    )
    $color = $colors[$Type]
    Write-Host $Message -ForegroundColor $color
}

function Test-Requirements {
    param([string]$RequirementsFile)
    
    if (-not (Test-Path $RequirementsFile)) {
        Write-Log "  ⚠ requirements.txt nao encontrado" 'warning'
        return $false
    }
    return $true
}

function Find-Python {
    <#
    Procura por Python no sistema, tentando várias opções
    Retorna o comando Python ou $null se não encontrar
    #>
    
    # Tenta py.exe (Python Launcher para Windows)
    try {
        $output = & py -3 --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  ✓ Python encontrado via py.exe" 'success'
            return 'py -3'
        }
    } catch { }
    
    # Tenta python no PATH
    try {
        $output = & python --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  ✓ Python encontrado no PATH" 'success'
            return 'python'
        }
    } catch { }
    
    # Tenta python3 no PATH
    try {
        $output = & python3 --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  ✓ Python encontrado (python3) no PATH" 'success'
            return 'python3'
        }
    } catch { }
    
    # Se nenhuma opção funcionou
    Write-Log "✗ ERRO: Python nao foi encontrado no sistema!" 'error'
    Write-Log ""
    Write-Log "Este computador nao possui Python instalado ou configurado corretamente." 'warning'
    Write-Log ""
    Write-Log "SOLUCAO:" 'warning'
    Write-Log "1. Baixe Python em: https://www.python.org/downloads/" 'info'
    Write-Log "2. Execute o instalador" 'info'
    Write-Log "3. IMPORTANTE: Marque a opcao 'Add Python to PATH' durante instalacao" 'warning'
    Write-Log "4. Reinicie o computador" 'info'
    Write-Log "5. Execute novamente este arquivo" 'info'
    Write-Log ""
    Write-Log "Ou, se ja tem Python instalado mas nao no PATH:" 'warning'
    Write-Log "1. Localize onde Python esta instalado" 'info'
    Write-Log "2. Adicione o diretorio ao PATH do sistema (Propriedades do Sistema)" 'info'
    Write-Log ""
    
    return $null
}

# ============================================================
# INÍCIO DO SCRIPT
# ============================================================
Clear-Host
Write-Log "============================================================" 'header'
Write-Log "GERADOR DE PLANILHA DE MDFs - LAUNCHER" 'header'
Write-Log "============================================================" 'header'
Write-Log ""

# Determina caminhos
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ProjectDir ".venv"
$VenvScripts = Join-Path $VenvDir "Scripts"
$PythonExe = Join-Path $VenvScripts "python.exe"
$PipExe = Join-Path $VenvScripts "pip.exe"
$RequirementsFile = Join-Path $ProjectDir "requirements.txt"
$MainScript = Join-Path $ProjectDir "gerar_planilha.py"

# ============================================================
# PASSO 0: Encontrar Python
# ============================================================
Write-Log "[0/3] Procurando Python..." 'info'

$PythonCmd = Find-Python

if ($null -eq $PythonCmd) {
    Read-Host "Pressione ENTER para sair"
    exit 1
}

Write-Log ""

# ============================================================
# PASSO 1: Verificar e criar Virtual Environment
# ============================================================
Write-Log "[1/3] Verificando ambiente virtual..." 'info'

if (-not (Test-Path $VenvDir)) {
    Write-Log "  → Criando novo ambiente virtual..." 'warning'
    
    try {
        Invoke-Expression "$PythonCmd -m venv `"$VenvDir`""
        if ($LASTEXITCODE -ne 0) {
            Write-Log "  ✗ Erro ao criar ambiente virtual" 'error'
            Read-Host "Pressione ENTER para sair"
            exit 1
        }
        Write-Log "  ✓ Ambiente virtual criado com sucesso" 'success'
    }
    catch {
        Write-Log "  ✗ Erro: $_" 'error'
        Read-Host "Pressione ENTER para sair"
        exit 1
    }
} else {
    Write-Log "  ✓ Ambiente virtual encontrado" 'success'
}

Write-Log ""

# ============================================================
# PASSO 2: Verificar e instalar dependências
# ============================================================
Write-Log "[2/3] Verificando dependencias..." 'info'

if (Test-Requirements $RequirementsFile) {
    Write-Log "  → Instalando pacotes de requirements.txt..." 'warning'
    
    try {
        & $PipExe install -q -r $RequirementsFile 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  ✓ Dependencias instaladas com sucesso" 'success'
        } else {
            Write-Log "  ⚠ Alguns pacotes podem estar com problema" 'warning'
        }
    }
    catch {
        Write-Log "  ⚠ Erro ao instalar dependências: $_" 'warning'
    }
} else {
    Write-Log "  → Instalando dependencias padrao..." 'warning'
    
    try {
        $packages = @('pandas', 'pdfplumber', 'openpyxl')
        & $PipExe install -q $packages 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Log "  ✓ Dependencias padrao instaladas" 'success'
        } else {
            Write-Log "  ⚠ Alguns pacotes podem estar com problema" 'warning'
        }
    }
    catch {
        Write-Log "  ⚠ Erro ao instalar dependências: $_" 'warning'
    }
}

# ============================================================
# PASSO 3: Executar programa principal
# ============================================================
Write-Log "[3/3] Iniciando programa..." 'info'
Write-Log ""

if (-not (Test-Path $MainScript)) {
    Write-Log "✗ ERRO: Script principal nao encontrado: $MainScript" 'error'
    Read-Host "Pressione ENTER para sair"
    exit 1
}

try {
    & $PythonExe $MainScript
}
catch {
    Write-Log "✗ Erro ao executar programa: $_" 'error'
    Read-Host "Pressione ENTER para sair"
    exit 1
}

Write-Log ""
Read-Host "Programa finalizado. Pressione ENTER para fechar"
