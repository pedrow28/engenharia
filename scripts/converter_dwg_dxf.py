"""
converter_dwg_dxf.py
====================
Converte arquivos DWG para DXF usando o ODA File Converter.

Uso:
    python converter_dwg_dxf.py <pasta_com_dwg>
    python converter_dwg_dxf.py <arquivo.dwg>

Dependência:
    ODA File Converter instalado (gratuito)
    https://www.opendesign.com/guestfiles/oda_file_converter
"""

import subprocess
import sys
import os
from pathlib import Path


def encontrar_oda_converter():
    """Procura o ODA File Converter no sistema."""
    possiveis = [
        Path(r"C:\Program Files\ODA"),
        Path(r"C:\Program Files (x86)\ODA"),
    ]
    
    for base in possiveis:
        if base.exists():
            for pasta in sorted(base.iterdir(), reverse=True):
                exe = pasta / "ODAFileConverter.exe"
                if exe.exists():
                    return str(exe)
    
    return None


def converter_dwg_para_dxf(caminho_entrada, caminho_saida=None, versao="ACAD2018"):
    """
    Converte DWG para DXF.
    
    Args:
        caminho_entrada: Pasta com DWGs ou arquivo DWG individual
        caminho_saida: Pasta de saída (padrão: mesma pasta)
        versao: Versão do AutoCAD (padrão: ACAD2018)
    """
    oda_path = encontrar_oda_converter()
    if not oda_path:
        print("ERRO: ODA File Converter não encontrado!")
        print("Baixe em: https://www.opendesign.com/guestfiles/oda_file_converter")
        return False
    
    print(f"ODA File Converter: {oda_path}")
    
    entrada = Path(caminho_entrada)
    
    if entrada.is_file():
        pasta_entrada = str(entrada.parent)
        pasta_saida = caminho_saida or pasta_entrada
        filtro = entrada.name
    elif entrada.is_dir():
        pasta_entrada = str(entrada)
        pasta_saida = caminho_saida or pasta_entrada
        filtro = "*.DWG"
    else:
        print(f"ERRO: Caminho não encontrado: {caminho_entrada}")
        return False
    
    os.makedirs(pasta_saida, exist_ok=True)
    
    cmd = [
        oda_path,
        pasta_entrada,
        pasta_saida,
        versao,
        "DXF",     # formato saída
        "0",       # sem recursão
        "1",       # com auditoria
        filtro,
    ]
    
    print(f"Convertendo: {pasta_entrada}")
    print(f"Saída: {pasta_saida}")
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"AVISO: Conversor retornou código {result.returncode}")
        if result.stderr:
            print(f"  {result.stderr}")
    
    # Verificar arquivos gerados
    dxf_files = list(Path(pasta_saida).glob('*.dxf'))
    print(f"\nArquivos DXF gerados: {len(dxf_files)}")
    for f in dxf_files:
        print(f"  {f.name}")
    
    return len(dxf_files) > 0


def main():
    if len(sys.argv) < 2:
        print("Uso: python converter_dwg_dxf.py <pasta_com_dwg ou arquivo.dwg>")
        print("")
        print("  Converte DWG para DXF usando ODA File Converter.")
        sys.exit(1)
    
    caminho = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else None
    
    sucesso = converter_dwg_para_dxf(caminho, saida)
    
    if sucesso:
        print("\nConversão concluída!")
    else:
        print("\nConversão falhou!")
        sys.exit(1)


if __name__ == '__main__':
    main()
