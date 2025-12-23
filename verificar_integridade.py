#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de verificação de integridade do ambiente
Verifica se todas as dependências estão instaladas corretamente
"""

import sys
import subprocess
import os
from pathlib import Path

def check_python_version():
    """Verifica se a versão do Python é compatível"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 7):
        print(f"✗ Python {version.major}.{version.minor} detectado")
        print("✓ Python 3.7+ é requerido")
        return False
    print(f"✓ Python {version.major}.{version.minor} compatível")
    return True

def check_package(package_name, import_name=None):
    """Verifica se um pacote está instalado"""
    if import_name is None:
        import_name = package_name
    
    try:
        __import__(import_name)
        print(f"  ✓ {package_name}")
        return True
    except ImportError:
        print(f"  ✗ {package_name} não instalado")
        return False

def check_files():
    """Verifica se os arquivos necessários existem"""
    required_files = [
        'BASE.csv',
        'gerar_planilha.py',
    ]
    
    project_dir = Path(__file__).parent
    missing = []
    
    for file in required_files:
        file_path = project_dir / file
        if not file_path.exists():
            missing.append(file)
            print(f"  ✗ {file} não encontrado")
        else:
            print(f"  ✓ {file}")
    
    # Verificar planilha de escala
    escala_found = False
    for file in project_dir.glob('escala*.xlsx'):
        print(f"  ✓ Planilha de escala: {file.name}")
        escala_found = True
        break
    
    if not escala_found:
        print(f"  ⚠ Nenhuma planilha de escala encontrada (procurando 'escala*.xlsx')")
    
    return len(missing) == 0

def check_folders():
    """Verifica se as pastas necessárias existem"""
    project_dir = Path(__file__).parent
    required_folders = [
        'MDFs geradas',
        'MDFs geradas/SOROCABA',
        'MDFs geradas/ITU',
        'MDFs geradas/OUTRAS ORI-DES',
    ]
    
    missing = []
    for folder in required_folders:
        folder_path = project_dir / folder
        if not folder_path.exists():
            missing.append(folder)
            print(f"  ✗ {folder}/ não encontrada")
        else:
            print(f"  ✓ {folder}/")
    
    return len(missing) == 0

def main():
    print("="*60)
    print("VERIFICADOR DE INTEGRIDADE DO AMBIENTE")
    print("="*60)
    print()
    
    print("1. Versão do Python:")
    python_ok = check_python_version()
    print()
    
    print("2. Pacotes obrigatórios:")
    packages = [
        ('pandas', 'pandas'),
        ('pdfplumber', 'pdfplumber'),
        ('openpyxl', 'openpyxl'),
        ('tkinter', 'tkinter'),
    ]
    
    all_packages_ok = True
    for package, import_name in packages:
        if not check_package(package, import_name):
            all_packages_ok = False
    print()
    
    print("3. Arquivos necessários:")
    files_ok = check_files()
    print()
    
    print("4. Pastas necessárias:")
    folders_ok = check_folders()
    print()
    
    print("="*60)
    if python_ok and all_packages_ok and files_ok and folders_ok:
        print("✓ AMBIENTE OK - Pronto para executar!")
        return 0
    else:
        print("✗ PROBLEMAS DETECTADOS - Veja acima para detalhes")
        return 1

if __name__ == '__main__':
    sys.exit(main())
