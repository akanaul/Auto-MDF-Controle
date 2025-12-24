"""
Sistema de Geração Automática de Planilha de MDFs
Processa PDFs de MDFs e gera planilhas CSV/Excel com informações estruturadas
"""

import os
import pandas as pd
from datetime import datetime, timedelta
import re
import pdfplumber
import tkinter as tk
from tkinter import simpledialog, messagebox
import glob
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# ============================================================================
# CONFIGURAÇÃO INICIAL
# ============================================================================

root = tk.Tk()
root.withdraw()

print("="*60)
print("GERADOR DE PLANILHA DE MDFs")
print("="*60)


# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def solicitar_responsavel():
    """Solicita e valida o nome do responsável pela emissão"""
    while True:
        responsavel = simpledialog.askstring(
            "Responsavel pela Emissao",
            "Digite o nome do responsavel pela emissao:\n(apenas letras, sem numeros)"
        )
        
        # Se usuário cancelou
        if responsavel is None:
            responsavel = "NAO INFORMADO"
            messagebox.showwarning(
                "Aviso", 
                f"Responsavel nao informado, usando: {responsavel}"
            )
            return responsavel
        
        responsavel = responsavel.strip()
        
        # Validar se vazio
        if not responsavel:
            messagebox.showerror(
                "Erro", 
                "Campo obrigatorio. Digite o nome do responsavel"
            )
            continue
        
        # Validar caracteres (apenas letras, acentos e espaços)
        caracteres_validos = 'áàâãéèêíïóôõöùûüçñÁÀÂÃÉÈÊÍÏÓÔÕÖÙÛÜÇÑ'
        tem_numero = any(c.isdigit() for c in responsavel)
        caracteres_ok = all(
            c.isalpha() or c.isspace() or c in caracteres_validos 
            for c in responsavel
        )
        
        if tem_numero or not caracteres_ok:
            messagebox.showerror(
                "Erro", 
                "Digite apenas letras (sem numeros ou caracteres especiais)"
            )
            continue
        
        responsavel = responsavel.upper()
        messagebox.showinfo(
            "Sucesso", 
            f"Responsavel definido: {responsavel}"
        )
        return responsavel


def encontrar_arquivo_escala(base_path):
    """Procura arquivo de escala com prefixo 'ESCALA' e extensão .xlsx"""
    for arquivo in os.listdir(base_path):
        if (arquivo.lower().startswith('escala') and 
            arquivo.lower().endswith('.xlsx')):
            return os.path.join(base_path, arquivo)
    return None


def remover_arquivos_antigos(pasta, padrao):
    """Remove arquivos antigos mantendo apenas o mais recente"""
    for arquivo in glob.glob(os.path.join(pasta, padrao)):
        try:
            os.remove(arquivo)
            print(f"Arquivo antigo removido: {os.path.basename(arquivo)}")
        except Exception as e:
            print(f"Erro ao remover {arquivo}: {e}")


def extrair_dt_do_pdf(caminho_pdf):
    """Extrai o número de DT procurando por padrão 'DT:' seguido de número"""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    match = re.search(
                        r'DT:\s*["\']?(\d+)["\']?', 
                        texto, 
                        re.IGNORECASE
                    )
                    if match:
                        return match.group(1).strip()
        return ''
    except Exception as e:
        print(f"  Erro ao ler PDF {caminho_pdf}: {e}")
        return ''


def extrair_cte_do_pdf(caminho_pdf):
    """Extrai o número de CTE procurando por padrão 'CTE:' seguido de número"""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    match = re.search(
                        r'CTE:\s*["\']?(\d+)["\']?', 
                        texto, 
                        re.IGNORECASE
                    )
                    if match:
                        return match.group(1).strip()
        return ''
    except Exception as e:
        print(f"  Erro ao ler PDF {caminho_pdf}: {e}")
        return ''


def extrair_mdfe_do_pdf(caminho_pdf):
    """Extrai o número de MDFE do PDF"""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    # Tenta padrão 1: "Modelo Série Número"
                    match = re.search(
                        r'Modelo\s+Série\s+Número.*?\n.*?(\d{6})', 
                        texto, 
                        re.IGNORECASE | re.DOTALL
                    )
                    if match:
                        return match.group(1).strip()
                    
                    # Tenta padrão 2: "Número:"
                    match = re.search(
                        r'Número[:\s]+(\d{6})', 
                        texto, 
                        re.IGNORECASE
                    )
                    if match:
                        return match.group(1).strip()
        return ''
    except Exception as e:
        print(f"  Erro ao ler PDF {caminho_pdf}: {e}")
        return ''


def extrair_hora_mdfe_do_pdf(caminho_pdf):
    """Extrai a hora de emissão do MDFE (formato HH:MM:SS)"""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    match = re.search(
                        r'\d{2}/\d{2}/\d{4}\s+(\d{2}:\d{2}:\d{2})', 
                        texto
                    )
                    if match:
                        return match.group(1).strip()
        return ''
    except Exception as e:
        print(f"  Erro ao ler PDF {caminho_pdf}: {e}")
        return ''


def extrair_carreta_cavalo_do_pdf(caminho_pdf):
    """Extrai placas de carreta e cavalo (primeira e segunda placa)"""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    linhas = texto.split('\n')
                    
                    # Procura pela seção de "Placa RNTRC CPF Nome"
                    for i, linha in enumerate(linhas):
                        if 'Placa' in linha and 'RNTRC' in linha:
                            carreta = ''
                            cavalo = ''
                            
                            if i + 1 < len(linhas):
                                primeira_placa = linhas[i + 1].strip().split()
                                carreta = primeira_placa[0] if primeira_placa else ''
                            
                            if i + 2 < len(linhas):
                                segunda_placa = linhas[i + 2].strip().split()
                                cavalo = segunda_placa[0] if segunda_placa else ''
                            
                            return carreta, cavalo
        return '', ''
    except Exception as e:
        print(f"  Erro ao ler PDF {caminho_pdf}: {e}")
        return '', ''


def extrair_nf_do_pdf(caminho_pdf):
    """Extrai número de NF (pode ser simples ou múltiplo: 280904/280894)"""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    match = re.search(
                        r'NF:\s*(\d+(?:/\d+)*)', 
                        texto, 
                        re.IGNORECASE
                    )
                    if match:
                        return match.group(1).strip()
        return ''
    except Exception as e:
        print(f"  Erro ao ler PDF {caminho_pdf}: {e}")
        return ''


def processar_pdfs_extrair_dados(pdfs, pdf_localizacao):
    """Processa todos os PDFs e extrai dados estruturados"""
    dados = {
        'dt': {},
        'cte': {},
        'mdfe': {},
        'hora_mdfe': {},
        'carreta': {},
        'cavalo': {},
        'nf': {}
    }
    
    print("\nExtraindo dados dos PDFs...")
    
    for pdf_nome in pdfs:
        pdf_limpo = re.sub(r'\s*\(.*?\)', '', pdf_nome).strip().upper()
        
        if pdf_limpo not in pdf_localizacao:
            continue
        
        subfolder, caminho_pdf = pdf_localizacao[pdf_limpo]
        
        # Extrai DT
        dt = extrair_dt_do_pdf(caminho_pdf)
        dados['dt'][pdf_limpo] = dt
        if dt:
            print(f"  {pdf_nome} [{subfolder}] -> DT: {dt}")
        
        # Extrai CTE
        cte = extrair_cte_do_pdf(caminho_pdf)
        dados['cte'][pdf_limpo] = cte
        if cte:
            print(f"  {pdf_nome} [{subfolder}] -> CTE: {cte}")
        
        # Extrai MDFE
        mdfe = extrair_mdfe_do_pdf(caminho_pdf)
        dados['mdfe'][pdf_limpo] = mdfe
        if mdfe:
            print(f"  {pdf_nome} [{subfolder}] -> MDFE: {mdfe}")
        
        # Extrai Hora MDFE
        hora_mdfe = extrair_hora_mdfe_do_pdf(caminho_pdf)
        dados['hora_mdfe'][pdf_limpo] = hora_mdfe
        if hora_mdfe:
            print(f"  {pdf_nome} [{subfolder}] -> Hora: {hora_mdfe}")
        
        # Extrai Carreta e Cavalo
        carreta, cavalo = extrair_carreta_cavalo_do_pdf(caminho_pdf)
        dados['carreta'][pdf_limpo] = carreta
        dados['cavalo'][pdf_limpo] = cavalo
        if carreta:
            print(f"  {pdf_nome} [{subfolder}] -> Carreta: {carreta}")
        
        # Extrai NF
        nf = extrair_nf_do_pdf(caminho_pdf)
        dados['nf'][pdf_limpo] = nf
        if nf:
            print(f"  {pdf_nome} [{subfolder}] -> NF: {nf}")
    
    return dados


# ============================================================================
# PROGRAMA PRINCIPAL
# ============================================================================

# Obter responsável
responsavel = solicitar_responsavel()

# Fechar janela
root.destroy()

print(f"Responsavel: {responsavel}\n")
print("Iniciando processamento...")
print("="*60 + "\n")


# Configurações de caminhos
BASE_PATH = str(Path(__file__).parent)
PDF_BASE_FOLDER = os.path.join(BASE_PATH, 'MDFs geradas')
PDF_SUBFOLDERS = ['SOROCABA', 'ITU', 'OUTRAS ORI-DES']
BASE_CSV = os.path.join(BASE_PATH, 'BASE.csv')
CSV_FOLDER = os.path.join(BASE_PATH, 'CSV')
EXCEL_FOLDER = os.path.join(BASE_PATH, 'EXCEL')


# Encontrar arquivo de escala
print("Buscando arquivo de escala...")
EXCEL_FILE = encontrar_arquivo_escala(BASE_PATH)

if EXCEL_FILE is None:
    EXCEL_FILE = os.path.join(BASE_PATH, 'ESCALA MOTORISTAS 2025.xlsx')
    print(f"Aviso: Usando arquivo padrão: {EXCEL_FILE}\n")
else:
    print(f"Encontrado: {os.path.basename(EXCEL_FILE)}\n")


# Determinar data do arquivo
data_agora = datetime.now()
usar_proxima_data = data_agora.hour >= 22

if usar_proxima_data:
    data_arquivo = (data_agora + timedelta(days=1)).strftime('%d.%m.%Y')
    print(f"Horário: {data_agora.strftime('%H:%M')} - Usando data do dia seguinte")
else:
    data_arquivo = data_agora.strftime('%d.%m.%Y')
    print(f"Horário: {data_agora.strftime('%H:%M')} - Usando data de hoje")


# Definir nomes dos arquivos
csv_filename = f'PLANILHA MDFS {data_arquivo}.csv'
excel_filename = f'PLANILHA MDFS {data_arquivo}.xlsx'

csv_path_raiz = os.path.join(BASE_PATH, csv_filename)
excel_path_raiz = os.path.join(BASE_PATH, excel_filename)

csv_path_historico = os.path.join(CSV_FOLDER, csv_filename)
excel_path_historico = os.path.join(EXCEL_FOLDER, excel_filename)


# Limpar arquivos antigos
print("\nLimpando arquivos antigos...")
remover_arquivos_antigos(BASE_PATH, 'PLANILHA MDFS *.csv')
remover_arquivos_antigos(BASE_PATH, 'PLANILHA MDFS *.xlsx')


# Carregar dados base
print("\nCarregando dados base...")
import unicodedata

def _try_read_base_headers(path):
    encodings = ['utf-8', 'latin-1', 'cp1252']
    last_exc = None
    for enc in encodings:
        try:
            df_base = pd.read_csv(path, encoding=enc, nrows=0)
            cols = df_base.columns.tolist()
            # Remover BOM se presente
            cols = [str(c).lstrip('\ufeff') for c in cols]
            print(f"[OK] BASE.csv lido com encoding: {enc}")
            return cols, enc
        except Exception as e:
            last_exc = e
    print(f"Erro ao ler BASE.csv com encodings testados: {last_exc}")
    return [], None

colunas_base, _base_encoding = _try_read_base_headers(BASE_CSV)

# Normalização utilitária para comparação tolerante de nomes de colunas
def _normalize(s):
    if s is None:
        return ''
    s = str(s)
    s = s.strip()
    # Remover acentos e caracteres especiais
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return s.upper()

try:
    # Carregar a planilha com data_only=True para garantir que valores calculados por fórmulas sejam lidos (copiando valores, não fórmulas)
    wb_escala = load_workbook(EXCEL_FILE, data_only=True)
    ws_escala = wb_escala.active
    data_iter = ws_escala.values
    headers = next(data_iter)
    # Normalizar cabeçalhos para string
    headers = [str(h).strip() if h is not None else '' for h in headers]
    df_motoristas = pd.DataFrame(data_iter, columns=headers)
except Exception as e:
    print(f"Erro ao ler arquivo de escala: {e}")
    df_motoristas = pd.DataFrame()


# Processar motoristas
print("\nProcessando motoristas...")
motoristas_lista = []
motoristas_escala = {}
motoristas_frota = {}
motoristas_nome_completo = {}
motoristas_gpid = {}
motoristas_cpf = {}

# Construir mapa de cabeçalhos normalizados -> nome real
col_map = { _normalize(c): c for c in df_motoristas.columns }

def _get_row_val(row, desired_col):
    key = _normalize(desired_col)
    actual = col_map.get(key)
    if actual is None:
        return ''
    try:
        v = row.get(actual, '')
        if pd.isna(v):
            return ''
        return str(v).strip()
    except Exception:
        return ''

# Função para encontrar motorista correspondente a um nome de PDF usando regras tolerantes
def _find_motorista_for(pdf_name):
    pdf_n = _normalize(pdf_name)
    # 1) exata
    for mot in motoristas_lista:
        if _normalize(mot) == pdf_n:
            return mot
    # 2) token / startswith / contains
    for mot in motoristas_lista:
        mot_n = _normalize(mot)
        tokens = mot_n.split()
        if pdf_n in tokens:
            return mot
        if mot_n.startswith(pdf_n):
            return mot
        if pdf_n in mot_n:
            return mot
    return None

for _, row in df_motoristas.iterrows():
    mot_raw = _get_row_val(row, 'MOTORISTA')
    if not mot_raw:
        continue
    mot_limpo = re.sub(r'\s*\(.*?\)', '', mot_raw).strip()
    if not mot_limpo:
        continue

    motoristas_lista.append(mot_limpo)
    motoristas_escala[mot_limpo] = _get_row_val(row, 'ESCALA')
    motoristas_frota[mot_limpo] = _get_row_val(row, 'FROTA')
    motoristas_nome_completo[mot_limpo] = _get_row_val(row, 'NOME COMPLETO')
    motoristas_gpid[mot_limpo] = _get_row_val(row, 'GPID')
    motoristas_cpf[mot_limpo] = _get_row_val(row, 'CPF')

print(f"Motoristas encontrados: {len(motoristas_lista)}")
if len(motoristas_lista) == 0:
    print("DEBUG: Cabeçalhos na planilha de escala:", df_motoristas.columns.tolist())
    print("DEBUG: Primeiras 5 linhas da planilha de escala:")
    for i, r in enumerate(df_motoristas.head(5).to_dict(orient='records')):
        print(f"  {i+1}: {r}")


# Encontrar PDFs
print("\nProcurando arquivos PDF...")
pdfs = []
pdf_localizacao = {}

for subfolder in PDF_SUBFOLDERS:
    pdf_folder = os.path.join(PDF_BASE_FOLDER, subfolder)
    
    if not os.path.exists(pdf_folder):
        print(f"  [{subfolder}] Pasta não existe")
        continue
    
    # Listar PDFs (ignorando .gitignore e .gitkeep)
    arquivos = [
        f.replace('.pdf', '').replace('.PDF', '')
        for f in os.listdir(pdf_folder)
        if (f.lower().endswith('.pdf') and 
            f.lower() not in ['.gitignore', '.gitkeep'])
    ]
    
    if arquivos:
        print(f"  [{subfolder}] {len(arquivos)} arquivo(s)")
        for arquivo in arquivos:
            pdfs.append(arquivo)
            nome_pdf = f"{arquivo}.pdf"
            pdf_path = os.path.join(pdf_folder, nome_pdf)
            pdf_localizacao[arquivo.upper()] = (subfolder, pdf_path)
    else:
        print(f"  [{subfolder}] Nenhum arquivo")

print(f"\nTotal de PDFs: {len(pdfs)}\n")


# Extrair dados dos PDFs
dados_pdfs = processar_pdfs_extrair_dados(pdfs, pdf_localizacao)


# Processar e criar nova planilha
print("\nProcessando dados...")
dados_novos = []

for pdf_nome in pdfs:
    pdf_limpo = re.sub(r'\s*\(.*?\)', '', pdf_nome).strip().upper()
    
    # Procurar motorista correspondente (regras tolerantes)
    motorista_encontrado = None
    pdf_norm = pdf_limpo

    # 1) igualdade normalizada
    for mot in motoristas_lista:
        if _normalize(mot) == _normalize(pdf_norm):
            motorista_encontrado = mot
            matched_by = 'exact'
            break

    # 2) token / prefix / contains
    if not motorista_encontrado:
        for mot in motoristas_lista:
            mot_n = _normalize(mot)
            pdf_n = _normalize(pdf_norm)
            tokens = mot_n.split()
            if pdf_n in tokens or mot_n.startswith(pdf_n) or pdf_n in mot_n:
                motorista_encontrado = mot
                matched_by = 'token/contains'
                break

    if not motorista_encontrado:
        print(f"X {pdf_limpo} NAO encontrado no Excel")
        continue
    else:
        # debug menor quando a correspondência foi por contains (menos segura)
        if 'matched_by' in locals() and matched_by != 'exact':
            print(f"[WARN] {pdf_limpo} casou com {motorista_encontrado} (metodo: {matched_by})")

    
    # Criar linha com dados estruturados
    linha = {col: '' for col in colunas_base}
    
    # Preencher informações básicas
    linha['DATA'] = data_arquivo
    linha['MOTORISTA'] = motorista_encontrado
    # Preencher nome completo, gpid e cpf (valores copiados da planilha de escala)
    linha['NOME COMPLETO'] = motoristas_nome_completo.get(motorista_encontrado, '')
    linha['GPID'] = motoristas_gpid.get(motorista_encontrado, '')
    linha['CPF'] = motoristas_cpf.get(motorista_encontrado, '')
    linha['HORA ESCALA (P2)'] = motoristas_escala.get(motorista_encontrado, '')
    linha['FROTA (P2)'] = motoristas_frota.get(motorista_encontrado, '')
    
    # Preencher informações do PDF
    linha['CARRETA (P2)'] = dados_pdfs['carreta'].get(pdf_limpo, '')
    linha['CAVALO (P2)'] = ''
    linha['DT'] = dados_pdfs['dt'].get(pdf_limpo, '')
    linha['CTE (P2)'] = dados_pdfs['cte'].get(pdf_limpo, '')
    linha['Nº MDFE (P2)'] = dados_pdfs['mdfe'].get(pdf_limpo, '')
    linha['HORA MDFE (P2)'] = dados_pdfs['hora_mdfe'].get(pdf_limpo, '')
    linha['NF (P2)'] = dados_pdfs['nf'].get(pdf_limpo, '')
    
    # Preencher dados do responsável
    linha['EMITO POR (P2)'] = responsavel
    linha['RESPONSAVEL P2'] = responsavel
    linha['STATUS (P2)'] = 'FATURADO'
    
    # Preencher origem/destino baseado no local
    if pdf_limpo in pdf_localizacao:
        subfolder, _ = pdf_localizacao[pdf_limpo]
        if subfolder in ['ITU', 'SOROCABA']:
            linha['ORIGEM (ESCALA)'] = subfolder
            linha['DESTINO (ESCALA)'] = 'DHL'
    
    dados_novos.append(linha)
    print(f"[OK] {pdf_limpo} → {motorista_encontrado}")


# Gerar arquivos
if dados_novos:
    print("\nGerando arquivos...")
    df_novo = pd.DataFrame(dados_novos)

    # Garantir que todas as colunas da BASE existam no DataFrame de saída
    missing_cols = [c for c in colunas_base if c not in df_novo.columns]
    if missing_cols:
        print(f"⚠️ Colunas ausentes no DataFrame gerado (serão criadas vazias): {missing_cols}")
        for c in missing_cols:
            df_novo[c] = ''

        # Tentar realinhar colunas por correspondência tolerante (normalizada)
        col_map = { _normalize(c): c for c in df_novo.columns }
        desired_cols = []
        for c in colunas_base:
            nc = _normalize(c)
            if nc in col_map:
                desired_cols.append(col_map[nc])
            else:
                desired_cols.append(c)

        # Reindex com as colunas esperadas da BASE (criadas vazias se necessário)
        df_novo = df_novo.reindex(columns=desired_cols)
    else:
        df_novo = df_novo[colunas_base]

    # Salvar CSV
    try:
        df_novo.to_csv(csv_path_raiz, index=False, encoding='latin-1')
        print(f"[OK] CSV criado: {csv_filename}")
    except Exception as e:
        print(f"[ERRO] CSV raiz: {e}")
    
    try:
        df_novo.to_csv(csv_path_historico, index=False, encoding='latin-1')
        print(f"[OK] CSV arquivado: CSV/{csv_filename}")
    except Exception as e:
        print(f"[ERRO] CSV histórico: {e}")
    
    # Salvar Excel
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Dados"
        
        # Headers
        for col_num, col_name in enumerate(df_novo.columns, 1):
            ws.cell(row=1, column=col_num, value=col_name)
        
        # Dados
        for row_num, row_data in enumerate(
            dataframe_to_rows(df_novo, index=False, header=False), 2
        ):
            for col_num, cell_value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col_num, value=cell_value)
        
        wb.save(excel_path_raiz)
        print(f"[OK] Excel criado: {excel_filename}")
        
        wb.save(excel_path_historico)
        print(f"[OK] Excel arquivado: EXCEL/{excel_filename}")
    except Exception as e:
        print(f"[ERRO] Excel: {e}")
    
    # Resumo
    print(f"\n{'='*60}")
    print(f"SUCESSO!")
    print(f"{'='*60}")
    print(f"Registros processados: {len(df_novo)}")
    print(f"Colunas: {len(df_novo.columns)}")
    
    try:
        messagebox.showinfo(
            "Sucesso! OK",
            f"Automação concluída com sucesso!\n\n"
            f"Registros: {len(df_novo)}\n\n"
            f"Arquivos:\n"
            f"  - {csv_filename}\n"
            f"  - {excel_filename}\n\n"
            f"Salvos em:\n"
            f"  - Raiz do programa\n"
            f"  - Pastas CSV/ e EXCEL/"
        )
    except Exception as e:
        print(f"[AVISO] Dialog não exibido: {e}")
else:
    print(f"\n{'='*60}")
    print("ERRO - Nenhum motorista foi encontrado!")
    print(f"{'='*60}")
    
    try:
        messagebox.showerror(
            "Erro! X",
            "Nenhum motorista foi encontrado.\n\n"
            "Verifique:\n"
            "  - Se existe arquivo começando com 'ESCALA'\n"
            "  - Se existem PDFs nas subpastas\n"
            "  - Se os nomes dos PDFs correspondem aos motoristas"
        )
    except Exception as e:
        print(f"[AVISO] Dialog não exibido: {e}")
