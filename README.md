# Gerador de Planilha de MDFs

## 🚀 Como Executar

### Opção 1: Batch File (Recomendado - Mais Simples)
Clique duas vezes em **`executar.bat`**

O arquivo .bat irá automaticamente:
- ✓ Procurar por Python no sistema (suporta computadores corporativos)
- ✓ Criar ambiente virtual (.venv) se não existir
- ✓ Instalar/verificar dependências
- ✓ Executar o programa

### Opção 2: PowerShell
Execute no PowerShell:
```powershell
.\executar.ps1
```

Nota: Pode ser necessário permitir execução de scripts. Se receber erro, execute:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Opção 3: Manual (Terminal)
```bash
.\.venv\Scripts\activate
python gerar_planilha.py
```

---

## 🆕 Novidades da Versão 2.0

- Compatibilidade corporativa reforçada: os launchers (batch e PowerShell) tentam `py.exe`, `python` e `python3` antes de falhar, e exibem instruções claras para recuperar o PATH faltante.
- `gerar_planilha.py` agora define `BASE_PATH` automaticamente com `Path(__file__).parent`, tornando o projeto resistente a renomeações ou movimentações de pasta.
- A documentação e as mensagens geradas destacam como lidar com computadores restritos, inclusive mostrando passos rápidos para configurar o PATH quando necessário.

---

## 🔴 Python não está instalado?

Se você está em um computador corporativo e vê a mensagem "Python não foi encontrado":

> Os launchers agora tentam `py.exe`, `python` e `python3` antes de exibir essa mensagem; use as soluções abaixo se mesmo assim o PATH continuar vazio.

**Solução 1: Instalar Python (Recomendado)**
1. Baixe Python em https://www.python.org/downloads/
2. Execute o instalador
3. **⚠️ IMPORTANTE:** Marque a opção **"Add Python to PATH"** durante a instalação
4. Reinicie o computador
5. Execute `executar.bat` novamente

**Solução 2: Usar Python Launcher (py.exe)**
Se Python já está instalado mas o PATH não está configurado:
- O programa tentará usar `py.exe` automaticamente
- Se ainda não funcionar, configure o PATH do Windows

**Solução 3: Adicionar ao PATH manualmente**
Se Python está instalado em um local específico:
1. Localize a pasta onde Python está instalado (ex: `C:\Python312`)
2. Abra "Propriedades do Sistema" → "Variáveis de Ambiente"
3. Edite a variável `Path` e adicione o caminho do Python
4. Reinicie o computador

---

## 📋 Arquivos Principais

- **`executar.bat`** - Launcher automático (Windows Batch)
- **`executar.ps1`** - Launcher avançado (PowerShell)
- **`gerar_planilha.py`** - Programa principal de automação
- **`requirements.txt`** - Lista de dependências Python
- **`BASE.csv`** - Template de base de dados
- **`ESCALA MOTORISTAS 2025.xlsx`** - Planilha de escala (detectada automaticamente)

---

## 🔧 Funcionamento

### Primeira Execução
1. Detecta que não existe `.venv`
2. Cria o ambiente virtual
3. Instala todas as dependências (pandas, pdfplumber, openpyxl)
4. Executa o programa

### Execuções Subsequentes
1. Valida a integridade do ambiente virtual
2. Verifica se todas as dependências estão instaladas
3. Executa o programa direto

---

## 📦 Dependências

O programa requer:
- **pandas** ≥ 1.0.0 - Manipulação de dados
- **pdfplumber** ≥ 0.7.0 - Extração de texto de PDFs
- **openpyxl** ≥ 3.0.0 - Geração de arquivos Excel

Todas instaladas automaticamente na primeira execução.

---

## 🎯 Funcionalidades

- ✅ Extração automática de dados de PDFs (DT, CTE, MDFE, NF, etc.)
- ✅ Busca automática de planilha de escala (qualquer nome começando com "ESCALA")
- ✅ Geração de CSV e Excel
- ✅ Arquivamento automático em pastas CSV/ e EXCEL/
- ✅ Limpeza de arquivos antigos na raiz
- ✅ Interface com diálogo GUI para entrada de dados
- ✅ Suporte a múltiplas subpastas de MDFs (SOROCABA, ITU, OUTRAS ORI-DES)
- ✅ Compatibilidade com computadores corporativos: `executar.bat` e `executar.ps1` tentam `py.exe`, `python` e `python3` antes de emitir orientações guiadas.
- ✅ Caminhos resilientes ao mover a pasta: `BASE_PATH` usa o diretório onde o script está sendo executado.

---

## 🧭 Regras de negócio principais

- SOROCABA: `MOTIVO ATRASO (P2)` fica "VETADO ANTECIPACAO DE MDF" e `HORA ESCALA (P2)` é mantido em branco.
- ITU: ao casar motorista, prioriza a 1ª aba visível da planilha (aba mais recente). Se não encontrar, procura em qualquer aba visível.
- ORIGEM/DESTINO: se o PDF vier de SOROCABA ou ITU, ORIGEM recebe a pasta e DESTINO recebe "DHL".
- Veículos: `CAVALO (P2)` e `CARRETA (P2)` são lidos do PDF; `FROTA (P2)` vem da planilha da escala.
- Identificação: nomes e aliases são normalizados (sem acento, maiúsculo) para casar PDF ↔ escala; remove conteúdo entre parênteses no nome.

---

## 🧱 Arquitetura do código

- Arquivo único: `gerar_planilha.py` concentra o fluxo.
- Etapas em funções puras:
	- leitura da escala (`carregar_escala`) e mapeamento de motoristas (`preparar_motoristas`)
	- varredura dos PDFs (`listar_pdfs`) e extrações por campo (`extrair_*`)
	- casamento PDF ↔ motorista (`_match_motorista`) com prioridade especial para ITU
	- montagem de linhas alinhadas ao cabeçalho do `BASE.csv` (`montar_registros`)
	- persistência (CSV/Excel) em `salvar_saidas`
- Helpers reutilizáveis: normalização de texto, progress bar, limpeza de arquivos antigos, cálculo de data.
- Execução centralizada em `main()`, chamada apenas quando o arquivo é executado diretamente.
- `BASE_PATH` é calculado como `Path(__file__).parent`, garantindo que os caminhos relativos funcionem mesmo que a pasta seja movida.

Para evoluir o código, prefira adicionar novas regras dentro de `montar_registros` ou novos extratores seguindo o padrão `extrair_*`.

---

## 📁 Estrutura de Pastas

```
Projeto automação MDF/
├── executar.bat              ← CLIQUE AQUI para executar
├── executar.ps1
├── gerar_planilha.py
├── requirements.txt
├── BASE.csv
├── ESCALA MOTORISTAS 2025.xlsx
├── .venv/                    ← Criado automaticamente
├── MDFs geradas/
│   ├── SOROCABA/
│   ├── ITU/
│   └── OUTRAS ORI-DES/
├── CSV/                      ← Histórico de CSVs
├── EXCEL/                    ← Histórico de Excels
└── PLANILHA MDFS DD.MM.YYYY.*   ← Arquivos gerados (últimos)
```

---

## ⚡ Dicas

1. **Primeira vez?** Apenas clique em `executar.bat` e deixe rodar
2. **Sem Python instalado?** Instale de https://www.python.org
3. **Problema com permissões?** Execute o terminal como administrador
4. **Adicionar nova dependência?** Edite `requirements.txt` e execute novamente

---

## 🐛 Troubleshooting

| Problema | Solução |
|----------|---------|
| "Python not found" | Instale Python 3.7+ de https://www.python.org (marque "Add to PATH") |
| Python instalado mas script não encontra | Tente usar Python Launcher ou configure PATH manualmente |
| Erro ao criar venv | Execute como administrador ou verifique permissões da pasta |
| Dependências não instalam | Verifique sua conexão de internet |
| Arquivo Excel não encontrado | Nomeie qualquer planilha começando com "ESCALA" |
| Computador corporativo - erro | Consulte seu TI para instalar Python com PATH configurado |

---

**Versão: 2.0 | Data: 03/01/2026**
