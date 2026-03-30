"""
Converte todos os arquivos DXF da pasta 'desenhos' para PDF.
Cada PDF é salvo ao lado do DXF original com o mesmo nome.
"""
import os
import sys
import glob
import traceback

import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend, layout
from ezdxf.addons.drawing.pymupdf import PyMuPdfBackend
from ezdxf.addons.drawing.config import Configuration


def convert_dxf_to_pdf(dxf_path: str, pdf_path: str) -> bool:
    """Converte um arquivo DXF para PDF. Retorna True se teve sucesso."""
    try:
        print(f"  Lendo: {dxf_path}")
        doc = ezdxf.readfile(dxf_path)
        msp = doc.modelspace()

        # Tamanho automático em mm
        page = layout.Page(0, 0, layout.Units.mm)
        backend = PyMuPdfBackend()
        config = Configuration()

        print("  Renderizando...")
        Frontend(RenderContext(doc), backend, config).draw_layout(msp, page)

        pdf_bytes = backend.get_pdf_bytes(page)
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)

        size_kb = os.path.getsize(pdf_path) / 1024
        print(f"  ✓ PDF salvo: {pdf_path} ({size_kb:.0f} KB)")
        return True

    except Exception as e:
        print(f"  ✗ ERRO ao converter {dxf_path}:")
        traceback.print_exc()
        return False


def main():
    # Define o diretório base como o pai do diretório 'scripts'
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    desenhos_dir = os.path.join(base_dir, "desenhos")

    if not os.path.isdir(desenhos_dir):
        print(f"Pasta 'desenhos' não encontrada em: {desenhos_dir}")
        sys.exit(1)

    # Busca todos os .dxf recursivamente
    dxf_files = glob.glob(os.path.join(desenhos_dir, "**", "*.dxf"), recursive=True)

    if not dxf_files:
        print("Nenhum arquivo .dxf encontrado!")
        sys.exit(1)

    print(f"Encontrados {len(dxf_files)} arquivo(s) DXF:")
    for f in dxf_files:
        print(f"  - {os.path.relpath(f, base_dir)}")
    print()

    sucesso = 0
    falha = 0

    for dxf_path in dxf_files:
        # PDF vai ser salvo no mesmo diretório do DXF
        pdf_name = os.path.splitext(os.path.basename(dxf_path))[0] + ".pdf"
        pdf_path = os.path.join(os.path.dirname(dxf_path), pdf_name)

        print(f"[{sucesso + falha + 1}/{len(dxf_files)}] Convertendo...")
        if convert_dxf_to_pdf(dxf_path, pdf_path):
            sucesso += 1
        else:
            falha += 1
        print()

    print("=" * 50)
    print(f"Concluído! {sucesso} sucesso(s), {falha} falha(s).")


if __name__ == "__main__":
    main()
