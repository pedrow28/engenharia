"""
converter_dwg_dxf.py
====================
Converte arquivos DWG para DXF usando o ODA File Converter.

Uso:
    python converter_dwg_dxf.py <pasta_com_dwg>
    python converter_dwg_dxf.py <arquivo.dwg>

Env:
    BATCH_SIZE=N   tamanho do lote (padrão 50; 0 desativa batching para pasta)

Dependência:
    ODA File Converter instalado (gratuito)
    https://www.opendesign.com/guestfiles/oda_file_converter

Comportamento em pasta (com batching):
- Resume automático: DWGs que já têm .dxf correspondente são ignorados.
- Cada lote é isolado em staging próprio (hardlinks — sem copiar arquivos).
- Se um lote falhar, os anteriores permanecem persistidos em disco. Rodar
  o script novamente reaproveita o que já foi feito.
"""

import subprocess
import sys
import os
import time
import shutil
import tempfile
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


# =============================================================================
# HELPERS DE BATCHING
# =============================================================================

def listar_dwgs_pendentes(pasta: Path):
    """
    Lista os DWGs da pasta separando em (todos, pendentes).

    Um DWG é "pendente" se ainda não existe um .dxf com mesmo nome base.
    É isso que dá resume automático: rodar o script de novo pula o que já
    foi convertido. O Windows tem FS case-insensitive, mas checamos ambos
    os casos por segurança.

    Helper público — importado pelo extractor no pipeline em lotes.
    """
    dwgs = sorted(
        (p for p in pasta.iterdir()
         if p.is_file() and p.suffix.lower() == '.dwg'),
        key=lambda p: p.name.lower(),
    )
    pendentes = []
    for dwg in dwgs:
        dxf_lower = dwg.with_suffix('.dxf')
        dxf_upper = dwg.with_suffix('.DXF')
        if not dxf_lower.exists() and not dxf_upper.exists():
            pendentes.append(dwg)
    return dwgs, pendentes


def converter_lote(oda_path: str, pasta_origem: Path, lote: list,
                   numero_lote: int, total_lotes: int, versao: str = "ACAD2018"):
    """
    Converte um lote de DWGs usando um staging temporário.

    Fluxo:
    1. Cria staging dentro de pasta_origem (mesmo drive → hardlink gratuito).
    2. Hardlinks os DWGs do lote para o staging (fallback: cópia).
    3. Chama ODA no staging (entrada = saída = staging).
    4. Move os .dxf gerados de volta para pasta_origem.
    5. Remove o staging (try/finally garante cleanup).

    Retorna o número de DXFs efetivamente persistidos na pasta original.

    Helper público — usado tanto pelo CLI standalone quanto pelo pipeline
    integrado do extractor.
    """
    t0 = time.perf_counter()
    print(f"\n[lote {numero_lote}/{total_lotes}] convertendo {len(lote)} arquivo(s)...",
          flush=True)

    try:
        staging = Path(tempfile.mkdtemp(prefix='_oda_batch_', dir=str(pasta_origem)))
    except OSError as e:
        print(f"  [lote {numero_lote}] ✗ não consegui criar staging: {e}", flush=True)
        return 0

    try:
        # Etapa 1: hardlink DWGs do lote para o staging
        linkados = 0
        for dwg in lote:
            link_path = staging / dwg.name
            try:
                os.link(str(dwg), str(link_path))
            except OSError:
                # Fallback: copy2 (drive diferente, permissão, FAT32…)
                try:
                    shutil.copy2(str(dwg), str(link_path))
                except Exception as e:
                    print(f"  [lote {numero_lote}] ⚠ falha ao preparar {dwg.name}: {e}",
                          flush=True)
                    continue
            linkados += 1

        if linkados == 0:
            print(f"  [lote {numero_lote}] ✗ nenhum arquivo preparado para conversão",
                  flush=True)
            return 0

        # Etapa 2: rodar ODA no staging (entrada e saída = staging)
        cmd = [
            oda_path,
            str(staging),
            str(staging),
            versao,
            "DXF",   # formato de saída
            "0",     # sem recursão
            "1",     # com auditoria
            "*.DWG",
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  [lote {numero_lote}] ⚠ ODA retornou código {result.returncode}",
                  flush=True)
            if result.stderr:
                stderr_short = result.stderr.strip().replace('\n', ' | ')[:200]
                print(f"  [lote {numero_lote}] stderr: {stderr_short}", flush=True)

        # Etapa 3: coletar DXFs gerados (dedup case-insensitive)
        dxfs_gerados = list(staging.glob('*.dxf')) + list(staging.glob('*.DXF'))
        dxfs_unicos = {d.name.lower(): d for d in dxfs_gerados}

        # Etapa 4: mover DXFs de volta para pasta original
        movidos = 0
        for dxf in dxfs_unicos.values():
            destino = pasta_origem / dxf.name
            try:
                if destino.exists():
                    destino.unlink()
                shutil.move(str(dxf), str(destino))
                movidos += 1
            except Exception as e:
                print(f"  [lote {numero_lote}] ⚠ falha ao mover {dxf.name}: {e}",
                      flush=True)

        dt = time.perf_counter() - t0
        status = "✓" if movidos == len(lote) else "⚠"
        media = dt / max(1, len(lote))
        print(f"  [lote {numero_lote}] {status} {movidos}/{len(lote)} DXFs em "
              f"{dt:.1f}s (média {media:.1f}s/arq)", flush=True)
        return movidos
    finally:
        # Cleanup: remove staging (hardlinks + quaisquer arquivos restantes)
        shutil.rmtree(str(staging), ignore_errors=True)


def _converter_arquivo_unico(oda_path: str, entrada: Path, caminho_saida, versao: str):
    """Modo arquivo-único: comportamento simples, sem batching/staging."""
    pasta_entrada = str(entrada.parent)
    pasta_saida = caminho_saida or pasta_entrada
    os.makedirs(pasta_saida, exist_ok=True)

    cmd = [
        oda_path,
        pasta_entrada,
        pasta_saida,
        versao,
        "DXF",
        "0",
        "1",
        entrada.name,
    ]
    print(f"Convertendo arquivo único: {entrada.name}", flush=True)
    print(f"Saída: {pasta_saida}", flush=True)

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"AVISO: Conversor retornou código {result.returncode}", flush=True)
        if result.stderr:
            print(f"  {result.stderr.strip()[:200]}", flush=True)

    dxf_esperado = entrada.with_suffix('.dxf')
    if dxf_esperado.exists():
        print(f"✓ {dxf_esperado.name}", flush=True)
        return True
    print(f"✗ DXF não gerado: {dxf_esperado.name}", flush=True)
    return False


# =============================================================================
# ENTRY POINT DA CONVERSÃO
# =============================================================================

def converter_dwg_para_dxf(caminho_entrada, caminho_saida=None, versao="ACAD2018"):
    """
    Converte DWG para DXF.

    - Arquivo único: chama ODA diretamente.
    - Pasta: processa em lotes (padrão BATCH_SIZE=50). Suporta resume
      automático — DWGs que já têm DXF correspondente são ignorados.

    Args:
        caminho_entrada: Pasta com DWGs ou arquivo DWG individual.
        caminho_saida: (só para arquivo único) pasta de saída; padrão = mesma pasta.
        versao: Versão do AutoCAD (padrão: ACAD2018).
    """
    oda_path = encontrar_oda_converter()
    if not oda_path:
        print("ERRO: ODA File Converter não encontrado!")
        print("Baixe em: https://www.opendesign.com/guestfiles/oda_file_converter")
        return False

    print(f"ODA File Converter: {oda_path}", flush=True)

    entrada = Path(caminho_entrada)

    if entrada.is_file():
        return _converter_arquivo_unico(oda_path, entrada, caminho_saida, versao)

    if not entrada.is_dir():
        print(f"ERRO: Caminho não encontrado: {caminho_entrada}")
        return False

    # Modo pasta: listagem + resume + batching
    todos_dwgs, pendentes = listar_dwgs_pendentes(entrada)

    if not todos_dwgs:
        print(f"Nenhum arquivo .dwg em: {entrada}")
        return False

    ja_feitos = len(todos_dwgs) - len(pendentes)
    print(f"DWGs na pasta: {len(todos_dwgs)}  |  já convertidos: {ja_feitos}  "
          f"|  pendentes: {len(pendentes)}", flush=True)

    if not pendentes:
        print("Nada a fazer — todos os DXFs já existem. (resume automático)",
              flush=True)
        return True

    # Batch size via env var — consistente com o extractor
    try:
        batch_size = int(os.environ.get('BATCH_SIZE', '50'))
    except ValueError:
        batch_size = 50

    if batch_size <= 0:
        # Batching desativado: um lote único com tudo
        batch_size = len(pendentes)

    n_lotes = (len(pendentes) + batch_size - 1) // batch_size

    print(f"\n{'=' * 56}", flush=True)
    print(f"  CONVERSÃO EM LOTES: {len(pendentes)} pendentes → "
          f"{n_lotes} lote(s) de até {batch_size}", flush=True)
    print(f"{'=' * 56}", flush=True)

    total_movidos = 0
    t_geral = time.perf_counter()

    for i in range(n_lotes):
        inicio = i * batch_size
        fim = min(inicio + batch_size, len(pendentes))
        lote = pendentes[inicio:fim]

        try:
            movidos = converter_lote(oda_path, entrada, lote, i + 1, n_lotes, versao)
            total_movidos += movidos
            if movidos < len(lote):
                faltam = len(lote) - movidos
                print(f"  [lote {i + 1}] ⚠ {faltam} arquivo(s) não foram convertidos "
                      f"neste lote.", flush=True)
        except KeyboardInterrupt:
            print(f"\n  ⚠ Interrompido — {total_movidos} arquivo(s) já persistidos.",
                  flush=True)
            print(f"  Rerun pulará os já feitos (resume automático).", flush=True)
            return total_movidos > 0
        except Exception as e:
            print(f"  [lote {i + 1}] ✗ ERRO: {e}", flush=True)
            print(f"  ⚠ {total_movidos} arquivo(s) já persistidos até aqui. "
                  f"Rerun pulará os já feitos.", flush=True)
            return total_movidos > 0

    dt = time.perf_counter() - t_geral
    media = dt / len(pendentes) if pendentes else 0.0
    print(f"\n[conversão] {total_movidos}/{len(pendentes)} DWGs convertidos em "
          f"{dt:.1f}s  |  média: {media:.1f}s/arquivo", flush=True)

    return total_movidos > 0


def main():
    if len(sys.argv) < 2:
        print("Uso: python converter_dwg_dxf.py <pasta_com_dwg ou arquivo.dwg>")
        print("")
        print("  Converte DWG para DXF usando ODA File Converter.")
        print("  Em modo pasta, processa em lotes (padrão 50) com resume automático.")
        print("")
        print("  Env: BATCH_SIZE=N (padrão 50; 0 desativa batching)")
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
