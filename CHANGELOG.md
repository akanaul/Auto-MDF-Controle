# Changelog - Auto Controle MDF

## Versão 2.0.0 - Melhorias de Compatibilidade

### ✅ Mudanças Implementadas

#### 1. **Compatibilidade com Computadores Corporativos**
- **Problema**: Computadores corporativos frequentemente não possuem Python no PATH
- **Solução**: 
  - `executar.bat`: Tenta localizar Python usando `py.exe`, `python` e `python3`
  - `executar.ps1`: Função `Find-Python()` que tenta múltiplas formas de encontrar Python
  - Mensagens de erro claras e instruções quando Python não é encontrado

#### 2. **Caminho Dinâmico para Pastas**
- **Problema**: Hardcoded path em `gerar_planilha.py` quebrava quando a pasta era renomeada
- **Solução**:
  - Alterado `BASE_PATH = r'c:\Users\Lenovo\Documents\Projeto automação MDF'`
  - Para: `BASE_PATH = str(Path(__file__).parent)` (diretório do script)
  - Importado `from pathlib import Path`
  - Agora o programa funciona independentemente do nome da pasta

#### 3. **Melhorias na Documentação**
- Atualizado `README.md` com seção "Python não está instalado?"
- Instruções específicas para:
  - Instalar Python com PATH
  - Usar Python Launcher (`py.exe`)
  - Configurar PATH manualmente
  - Ajuda para computadores corporativos

---

### 📝 Detalhes Técnicos

#### executar.bat
```batch
# Antes:
python -m venv "%VENV_DIR%"

# Depois:
REM Tenta py.exe, python e python3 em sequência
set PYTHON_CMD=
if errorlevel equ 0 (
    set PYTHON_CMD=py -3
    ...
)
```

#### executar.ps1
```powershell
# Nova função Find-Python que:
- Tenta py -3 (Python Launcher)
- Tenta python (PATH)
- Tenta python3 (PATH)
- Fornece instruções claras se não encontrar
```

#### gerar_planilha.py
```python
# Antes:
BASE_PATH = r'c:\Users\Lenovo\Documents\Projeto automação MDF'

# Depois:
from pathlib import Path
BASE_PATH = str(Path(__file__).parent)
```

---

### 🎯 Benefícios

1. **Resisilência**: Programa não quebra ao renomear a pasta
2. **Compatibilidade**: Funciona em computadores corporativos sem Python no PATH
3. **User-Friendly**: Mensagens claras ajudam usuários a resolver problemas
4. **Flexibilidade**: Múltiplas formas de localizar Python

---

### ⚠️ Requisitos

- Python 3.7+ (instalado em qualquer local do sistema)
- Windows 10 ou superior
- PowerShell 5.0+ (para `executar.ps1`)

---

### 📦 Arquivos Modificados

- `executar.bat` ✓
- `executar.ps1` ✓
- `gerar_planilha.py` ✓
- `README.md` ✓

---

### 🔄 Compatibilidade com Versões Anteriores

Todas as mudanças são retrocompatíveis. Ambientes existentes continuarão funcionando normalmente.

---

**Data**: 21-22 de Dezembro de 2025
**Versão Python**: 3.7+
**Status**: ✅ Pronto para Produção
