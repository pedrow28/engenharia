"""
extrair_dados_dxf.py
====================
Extrai dados de engenharia de arquivos DXF (convertidos de DWG)
e gera um relatório com os dados prontos para a planilha.

Uso:
    python extrair_dados_dxf.py <arquivo.dxf>
    python extrair_dados_dxf.py <pasta_com_dxf>

Dependências:
    pip install ezdxf
"""

import ezdxf
import sys
import os
import re
import json
from collections import defaultdict
from pathlib import Path


# =============================================================================
# CONFIGURAÇÃO
# =============================================================================

# Peso específico do concreto (kg/m³)
PESO_ESPECIFICO_CONCRETO = 2400


# =============================================================================
# EXTRAÇÃO DO NOME DO ARQUIVO
# =============================================================================

def parse_filename(filename):
    """
    Extrai dados do nome do arquivo.
    Padrão: PEÇA-LARGURAxALTURAxCOMPRIMENTO-FORMA E ARMAÇÃO-RXX
    Exemplo: P5-30x80x1408-FORMA E ARMAÇÃO-R01
    """
    name = Path(filename).stem
    # Remover sufixo de cópia do Windows ex: " (1)", " (2)"
    name = re.sub(r'\s*\(\d+\)$', '', name)

    # Padrão com seção: nome-LxAxC-tipo-revisao
    match = re.match(
        r'^(.+?)-(\d+[xX]\d+(?:[xX][\d,.]+)?)-(.+?)-(R\d+)$', name
    )
    if match:
        titulo = match.group(1)
        secao_raw = match.group(2)
        parts = re.split(r'[xX]', secao_raw)
        if len(parts) == 3:
            largura, altura, comprimento = parts
            secao = f"{largura}x{altura}"
            comprimento = comprimento.replace(',', '.')
        else:
            secao = secao_raw
            comprimento = None
        return {
            'nome_arquivo': name,
            'titulo_peca': titulo,
            'secao': secao,
            'comprimento_cm': float(comprimento) if comprimento else None,
        }

    # Padrão sem seção (ex: lajes L201=L301-FORMA E ARMAÇÃO-R00)
    match2 = re.match(r'^(.+?)-(FORMA E ARMAÇÃO|FORMA|ARMAÇÃO)-(R\d+)$', name)
    if match2:
        return {
            'nome_arquivo': name,
            'titulo_peca': match2.group(1),
            'secao': None,
            'comprimento_cm': None,
        }

    return {
        'nome_arquivo': name,
        'titulo_peca': name,
        'secao': None,
        'comprimento_cm': None,
    }


# =============================================================================
# EXTRAÇÃO DOS BLOCOS
# =============================================================================

def extract_notas(msp):
    """
    Extrai dados do bloco NOTAS.
    A=fck, B=Volume, C=Peso concreto, D=Peso peça, E=fcj (ou posição 5/7=G), I=Cobrimento

    O fcj pode estar em diferentes atributos dependendo do template:
    - Atributo 'E' (posição 5): padrão para pilares
    - Atributo 'G' (posição 7): encontrado nas vigas (20,3 MPa)
    Tenta cada um na ordem até encontrar valor válido.
    """
    for e in msp.query('INSERT'):
        if e.dxf.name == 'NOTAS':
            notas = {a.dxf.tag: a.dxf.text for a in e.attribs}
            def safe_float(key):
                val = notas.get(key, '-')
                if not val or str(val).strip() in ('-', ''):
                    return None
                try:
                    return float(str(val).replace(',', '.'))
                except (ValueError, TypeError):
                    return None
            def first_valid(*keys):
                for k in keys:
                    v = safe_float(k)
                    if v is not None:
                        return v
                return None
            return {
                'fck_mpa': safe_float('A'),
                'volume_concreto_m3': safe_float('B'),
                'peso_concreto_kgf': safe_float('C'),
                'peso_peca_kgf': safe_float('D'),
                # fcj pode estar em E (pos.5) ou G (pos.7) dependendo do template
                'fcj_mpa': first_valid('E', 'G', '5', '7'),
                'cobrimento_cm': safe_float('I'),
            }
    return {}


def extract_carimbo(msp):
    """
    Extrai dados do bloco CARIMBO.
    Atributos têm tag="X", usamos posição (índice) para identificar.
    """
    for e in msp.query('INSERT'):
        if e.dxf.name == 'CARIMBO':
            attribs = [a.dxf.text for a in e.attribs]
            if len(attribs) >= 16:
                nome_com_qtd = attribs[2]
                match = re.search(r'\((\d+)[xX]\)', nome_com_qtd)
                nome_limpo = re.sub(r'\(\d+[xX]\)', '', nome_com_qtd).strip()
                if match:
                    quantidade = int(match.group(1))
                elif '=' in nome_limpo:
                    quantidade = len(nome_limpo.split('='))
                else:
                    quantidade = None  # sem padrão explícito
                return {
                    'tipo_desenho': attribs[0],
                    'titulo_peca': nome_limpo,
                    'quantidade': quantidade,
                    'secao_comprimento': attribs[4],
                    'cliente': attribs[5],
                    'obra': attribs[6],
                    'data': attribs[12],
                    'folha': attribs[15],
                }
            break
    return {}


# =============================================================================
# EXTRAÇÃO DAS TABELAS DE AÇO (LAYER ListaF)
# =============================================================================

def _extrair_peso_cp190rb(listaf, y_groups):
    """
    Detecta aço protendido CP190RB no layer ListaF.

    A tabela de lista de ferros tem dois tipos de linha com CP190RB:
    - Linha de item individual: "N° | CP190RB | diam | qtd | comp | peso"
      (o primeiro token da linha é o número do item — um inteiro)
    - Linha de resumo: "CP190RB | diam | comp_total | peso_metro | peso_total"
      (CP190RB é o primeiro token significativo da linha)

    O peso correto é o ÚLTIMO valor da linha de resumo (já inclui fator de
    alongamento). As linhas de item individual são ignoradas pois já estão
    contabilizadas no resumo.

    Retorna o peso total do aço protendido, ou 0 se não encontrado.
    """
    peso_protendido = 0.0
    encontrados = []

    for y, items in y_groups.items():
        items_sorted = sorted(items, key=lambda i: i[0])

        # Verificar se CP190RB está nesta linha
        cp_indices = [i for i, (x, txt) in enumerate(items_sorted)
                      if 'CP190RB' in txt.upper()]
        if not cp_indices:
            continue

        primeiro_cp_idx = cp_indices[0]

        # Determinar se é linha de item: verifica se há um inteiro imediatamente
        # antes do CP190RB (= número de item da lista de ferros)
        e_linha_item = False
        if primeiro_cp_idx > 0:
            txt_antes = items_sorted[primeiro_cp_idx - 1][1].strip()
            # Limpar formatação MTEXT
            txt_limpo = re.sub(r'\{\\[^;]*;([^}]*)\}', r'\1', txt_antes).strip()
            try:
                int(txt_limpo)
                e_linha_item = True
            except ValueError:
                pass

        if e_linha_item:
            # Linha de item individual: pular (já somada no resumo)
            continue

        # Linha de resumo: pegar o ÚLTIMO valor numérico (peso total com fator)
        valores = []
        for j in range(primeiro_cp_idx + 1, len(items_sorted)):
            txt_j = items_sorted[j][1].strip()
            try:
                val = float(txt_j.replace(',', '.'))
                if val > 0:
                    valores.append(val)
            except ValueError:
                continue

        if valores:
            peso_linha = valores[-1]
            encontrados.append({'y': y, 'x': items_sorted[primeiro_cp_idx][0], 'valor': peso_linha})
            peso_protendido += peso_linha

    if encontrados:
        print(f"  [DEBUG PROTENDIDO] CP190RB resumo encontrado ({len(encontrados)}):")
        for i, e in enumerate(encontrados):
            print(f"    [{i}] valor={e['valor']:.2f}  x={e['x']:.0f}  y={e['y']:.0f}")
    else:
        print(f"  [DEBUG PROTENDIDO] CP190RB não encontrado")

    return peso_protendido


def extract_peso_total_aco(msp):
    """
    Extrai valores PESO TOTAL (kg) das tabelas RESUMO DE AÇO.
    Retorna pesos separados: frouxo (armação principal), consolo e protendido (CP190RB).

    Lógica do protendido:
    - Se CP190RB for encontrado no ListaF, seu peso = protendido
    - O PESO TOTAL da tabela principal já inclui frouxo + protendido
    - Então: peso_frouxo = peso_total_tabela - peso_protendido
    """
    # Coletar textos do layer ListaF
    listaf = []
    for e in msp.query('TEXT'):
        if e.dxf.layer == 'ListaF':
            listaf.append((e.dxf.text, e.dxf.insert.x, e.dxf.insert.y))
    for e in msp.query('MTEXT'):
        if e.dxf.layer == 'ListaF':
            listaf.append((e.text, e.dxf.insert.x, e.dxf.insert.y))

    if not listaf:
        return {}

    # Agrupar por Y (linhas da tabela)
    y_groups = defaultdict(list)
    for txt, x, y in listaf:
        y_groups[round(y, 0)].append((x, txt))

    # --- Extrair peso do aço protendido (CP190RB) ---
    peso_protendido = _extrair_peso_cp190rb(listaf, y_groups)

    # --- Encontrar linhas "PESO TOTAL (kg)" e extrair valor numérico ---
    peso_entries = []
    for y, items in y_groups.items():
        items_sorted = sorted(items, key=lambda i: i[0])

        # Verificar se tem "PESO TOTAL" nesta linha
        peso_indices = [i for i, (x, txt) in enumerate(items_sorted) if 'PESO TOTAL' in txt.upper()]

        if peso_indices:
            for i in peso_indices:
                peso_x = items_sorted[i][0]

                # Procurar o valor numérico imediatamente DEPOIS do label
                if i + 1 < len(items_sorted):
                    prox_x, prox_txt = items_sorted[i+1]
                    try:
                        val = float(prox_txt.replace(',', '.'))
                        if prox_x - peso_x < 1000:
                            peso_entries.append({'y': y, 'x_label': peso_x, 'valor': val})
                    except ValueError:
                        continue

    if not peso_entries:
        return {}

    # DEBUG: todos os PESO TOTAL encontrados
    print(f"  [DEBUG CONSOLO] PESO TOTAL encontrados ({len(peso_entries)}):")
    for i, pe in enumerate(peso_entries):
        print(f"    [{i}] valor={pe['valor']:.2f}  x={pe['x_label']:.0f}  y={pe['y']:.0f}")

    # --- Classificar tabelas: consolo vs armação principal ---
    # Cada bloco INSERT 'LISTA DE FERROS CONSOLO' tem atributo 'D' com multiplicidade
    # ex: '1x', '2x', '3x' — quando um consolo se repete na peça.
    consolo_block_info = []  # lista de (x, y, multiplicador)
    for e in msp.query('INSERT'):
        if e.dxf.name == 'LISTA DE FERROS CONSOLO':
            x = e.dxf.insert.x
            by = e.dxf.insert.y
            mult = 1
            for a in e.attribs:
                if a.dxf.tag == 'D':
                    m = re.match(r'(\d+)\s*[xX]', a.dxf.text.strip())
                    if m:
                        mult = int(m.group(1))
            consolo_block_info.append((x, by, mult))

    consolo_block_xs = [x for x, by, _ in consolo_block_info]
    print(f"  [DEBUG CONSOLO] Blocos INSERT 'LISTA DE FERROS CONSOLO': {len(consolo_block_info)} -> {[(f'x={x:.0f}', f'{m}x') for x, _, m in consolo_block_info]}")

    if consolo_block_xs:
        threshold = 1500

        def _dist_eucl(pe, bx, by):
            return ((pe['x_label'] - bx) ** 2 + (pe['y'] - by) ** 2) ** 0.5

        consolo = [e for e in peso_entries
                   if any(abs(e['x_label'] - bx) < threshold for bx in consolo_block_xs)]
        frouxo = [e for e in peso_entries if e not in consolo]
        print(f"  [DEBUG] Threshold={threshold}: consolo_entries={len(consolo)}, frouxo_entries={len(frouxo)}")

        # Fallback: quando todas as entradas foram para consolo (tabelas muito próximas em X),
        # usar distância euclidiana — a mais distante de qualquer bloco consolo = tabela principal.
        if not frouxo and len(consolo) > 1:
            print(f"  [DEBUG] Frouxo vazio — usando distância euclidiana para separar tabela principal.")
            def _min_dist(pe):
                return min(_dist_eucl(pe, bx, by) for bx, by, _ in consolo_block_info)
            sorted_by_dist = sorted(consolo, key=_min_dist, reverse=True)
            frouxo = [sorted_by_dist[0]]
            consolo = sorted_by_dist[1:]
            print(f"  [DEBUG] Após fallback euclidiano: frouxo={[e['valor'] for e in frouxo]}, consolo={[e['valor'] for e in consolo]}")

    elif len(peso_entries) == 1:
        frouxo = peso_entries
        consolo = []
    else:
        sorted_entries = sorted(peso_entries, key=lambda e: e['valor'], reverse=True)
        frouxo = [sorted_entries[0]]
        consolo = sorted_entries[1:]

    # Peso total da tabela principal (frouxo + protendido juntos)
    peso_total_tabela = max(e['valor'] for e in frouxo) if frouxo else None

    # Se há protendido, subtrair do total para obter o frouxo puro
    if peso_total_tabela is not None and peso_protendido > 0:
        peso_frouxo = peso_total_tabela - peso_protendido
        print(f"  [DEBUG PROTENDIDO] peso_total_tabela={peso_total_tabela:.2f} - protendido={peso_protendido:.2f} = frouxo={peso_frouxo:.2f}")
    else:
        peso_frouxo = peso_total_tabela

    # Consolo: somar com multiplicidade (atributo 'D' do bloco INSERT)
    # Para cada peso_entry do consolo, encontrar o bloco mais próximo e aplicar multiplicador
    def _mult_consolo(pe_x):
        candidatos = [(abs(pe_x - bx), mult) for bx, by, mult in consolo_block_info
                      if abs(pe_x - bx) < threshold]
        if candidatos:
            return min(candidatos, key=lambda t: t[0])[1]
        return 1

    n_consolos = len(consolo_block_info)
    if consolo:
        peso_consolo = sum(e['valor'] * _mult_consolo(e['x_label']) for e in consolo)
    else:
        peso_consolo = None

    print(f"  [DEBUG CONSOLO] frouxo={[e['valor'] for e in frouxo]}  consolo={[(e['valor'], _mult_consolo(e['x_label'])) for e in consolo]}")
    print(f"  [DEBUG CONSOLO] n_blocos={n_consolos} soma_consolo={peso_consolo}")

    return {
        'peso_aco_frouxo_kg': peso_frouxo,
        'peso_aco_consolo_kg': peso_consolo,
        'peso_aco_protendido_kg': peso_protendido,
        '_n_consolos': n_consolos,
    }


def extract_lista_ferros_detalhada(msp):
    """
    Extrai todos os itens da LISTA DE FERROS (layer ListaF) organizados por tabela.
    Retorna as linhas agrupadas por Y para referência.
    """
    listaf = []
    for e in msp.query('TEXT'):
        if e.dxf.layer == 'ListaF':
            listaf.append((e.dxf.text, round(e.dxf.insert.x, 1), round(e.dxf.insert.y, 1)))
    for e in msp.query('MTEXT'):
        if e.dxf.layer == 'ListaF':
            txt = re.sub(r'\{\\[^;]*;([^}]*)\}', r'\1', e.text)  # limpar formatação
            listaf.append((txt, round(e.dxf.insert.x, 1), round(e.dxf.insert.y, 1)))

    y_groups = defaultdict(list)
    for txt, x, y in listaf:
        y_groups[y].append((x, txt))

    rows = []
    for y in sorted(y_groups.keys(), reverse=True):
        items = sorted(y_groups[y], key=lambda i: i[0])
        row_text = ' | '.join(txt for _, txt in items)
        rows.append({'y': y, 'conteudo': row_text})
    return rows


# =============================================================================
# EXTRAÇÃO PRINCIPAL
# =============================================================================

def extrair_dados_completos(filepath):
    """Extrai todos os dados de um arquivo DXF."""
    try:
        doc = ezdxf.readfile(filepath)
    except IOError:
        print(f"  ERRO: Não foi possível ler: {filepath}")
        return None
    except ezdxf.DXFStructureError:
        print(f"  ERRO: DXF inválido: {filepath}")
        return None

    msp = doc.modelspace()
    filename = os.path.basename(filepath)

    # Extrair de cada fonte
    dados_nome = parse_filename(filename)
    dados_notas = extract_notas(msp)
    dados_carimbo = extract_carimbo(msp)
    dados_aco = extract_peso_total_aco(msp)

    # Determinar tipo de peça
    titulo = dados_carimbo.get('titulo_peca', dados_nome.get('titulo_peca', ''))
    if re.match(r'^(VG|VL|VR|VTI|VTA|V\d)', titulo):
        tipo_peca = 'VIGA'
    elif re.match(r'^L', titulo):
        tipo_peca = 'LAJE'
    elif re.match(r'^P', titulo):
        tipo_peca = 'PILAR'
    else:
        tipo_peca = 'OUTRO'

    # Valores base — peças igualadas (P11=P12 → 2) sempre têm prioridade
    titulo_arquivo = dados_nome.get('titulo_peca', '')
    if '=' in titulo:
        quantidade = len(titulo.split('='))
    elif '=' in titulo_arquivo:
        quantidade = len(titulo_arquivo.split('='))
    elif dados_carimbo.get('quantidade') is not None:
        quantidade = dados_carimbo['quantidade']
    else:
        quantidade = 1
    secao = dados_nome.get('secao') or dados_carimbo.get('secao_comprimento', '')
    comprimento = dados_nome.get('comprimento_cm')
    volume = dados_notas.get('volume_concreto_m3')
    fck = dados_notas.get('fck_mpa')
    fcj = dados_notas.get('fcj_mpa')

    # Pesos de aço
    peso_frouxo = dados_aco.get('peso_aco_frouxo_kg')
    peso_protendido = dados_aco.get('peso_aco_protendido_kg', 0)
    peso_consolo = dados_aco.get('peso_aco_consolo_kg')

    # Pilares nunca têm aço protendido — qualquer CP190RB detectado
    # (ex: grampos de ancoragem no consolo) é revertido para o frouxo.
    if tipo_peca == 'PILAR' and peso_protendido:
        print(f"  [PILAR] Revertendo CP190RB ({peso_protendido:.2f} kg) para frouxo — pilares não têm protendido.")
        peso_frouxo = (peso_frouxo or 0) + peso_protendido
        peso_protendido = 0

    # Peso total unitário = peso concreto + frouxo + protendido + consolo
    peso_concreto = volume * PESO_ESPECIFICO_CONCRETO if volume else None
    if peso_concreto is not None and peso_frouxo is not None:
        peso_total_unit = peso_concreto + peso_frouxo + (peso_protendido or 0) + (peso_consolo or 0)
    else:
        peso_total_unit = None

    # Taxas (kg/m³) — taxa frouxo inclui consolo (aço passivo total)
    aco_passivo = (peso_frouxo or 0) + (peso_consolo or 0)
    taxa_frouxo = aco_passivo / volume if (aco_passivo and volume) else None
    taxa_protendido = peso_protendido / volume if (peso_protendido and volume) else None

    # Totais (multiplicados pela quantidade)
    volume_total = volume * quantidade if volume else None
    peso_total_t = peso_total_unit * quantidade / 1000 if peso_total_unit else None
    peso_frouxo_t = peso_frouxo * quantidade / 1000 if peso_frouxo else None
    peso_protendido_t = peso_protendido * quantidade / 1000 if peso_protendido else None

    return {
        'A_tipo_peca': tipo_peca,
        'B_titulo_peca': titulo,
        'C_nome_desenho': dados_nome['nome_arquivo'],
        'D_quantidade': quantidade,
        'E_secao': secao,
        'F_comprimento_cm': comprimento,
        'G_volume_concreto_m3': volume,
        'H_volume_concreto_total_m3': volume_total,
        'I_fck_mpa': fck,
        'J_fcj_mpa': fcj,
        'K_peso_aco_frouxo_kg': peso_frouxo,
        'L_peso_aco_protendido_kg': peso_protendido,
        'M_peso_aco_consolo_kg': peso_consolo,
        'M_n_consolos': dados_aco.get('_n_consolos', 0),
        'N_peso_total_unitario_kg': peso_total_unit,
        'O_taxa_aco_frouxo_kg_m3': taxa_frouxo,
        'P_taxa_aco_protendido_kg_m3': taxa_protendido,
        'Q_peso_total_t': peso_total_t,
        'R_peso_total_aco_frouxo_t': peso_frouxo_t,
        'S_peso_total_aco_protendido_t': peso_protendido_t,
    }


# =============================================================================
# RELATÓRIO
# =============================================================================

def formatar_relatorio(dados):
    """Imprime relatório comparativo para validação."""
    if not dados:
        return "Nenhum dado extraído."

    campos = [
        ('A', 'Tipo de Peça', 'A_tipo_peca', ''),
        ('B', 'Título da Peça', 'B_titulo_peca', ''),
        ('C', 'Nome Desenho', 'C_nome_desenho', ''),
        ('D', 'Quantidade', 'D_quantidade', ''),
        ('E', 'Seção', 'E_secao', ''),
        ('F', 'Comprimento', 'F_comprimento_cm', 'cm'),
        ('G', 'Volume Concreto', 'G_volume_concreto_m3', 'm³'),
        ('H', 'Vol. Concreto Total', 'H_volume_concreto_total_m3', 'm³'),
        ('I', 'fck', 'I_fck_mpa', 'MPa'),
        ('J', 'fcj', 'J_fcj_mpa', 'MPa'),
        ('K', 'Peso Aço Frouxo', 'K_peso_aco_frouxo_kg', 'kg'),
        ('L', 'Peso Aço Protendido', 'L_peso_aco_protendido_kg', 'kg'),
        ('M', 'Peso Aço Consolo', 'M_peso_aco_consolo_kg', 'kg'),
        ('N', 'Peso Total Unitário', 'N_peso_total_unitario_kg', 'kg'),
        ('O', 'Taxa Aço Frouxo', 'O_taxa_aco_frouxo_kg_m3', 'kg/m³'),
        ('P', 'Taxa Aço Protendido', 'P_taxa_aco_protendido_kg_m3', 'kg/m³'),
        ('Q', 'Peso Total', 'Q_peso_total_t', 't'),
        ('R', 'Peso Total Frouxo', 'R_peso_total_aco_frouxo_t', 't'),
        ('S', 'Peso Total Protendido', 'S_peso_total_aco_protendido_t', 't'),
    ]

    lines = []
    lines.append("=" * 65)
    lines.append(f"  PEÇA: {dados['B_titulo_peca']}  |  {dados['C_nome_desenho']}")
    lines.append("=" * 65)
    for col, nome, chave, un in campos:
        val = dados.get(chave)
        if val is None:
            val_str = '-'
        elif isinstance(val, float):
            val_str = f"{val:,.4f}" if abs(val) < 1 else f"{val:,.2f}"
        else:
            val_str = str(val)
        extra = ''
        if chave == 'M_peso_aco_consolo_kg':
            n = dados.get('M_n_consolos', 0)
            if n > 0:
                extra = f'  ({n}x)'
        lines.append(f"  {col:2s} | {nome:25s} | {val_str} {un}{extra}")
    lines.append("")
    return "\n".join(lines)


# =============================================================================
# ATUALIZAÇÃO DA PLANILHA (Excel)
# =============================================================================

def atualizar_planilha(caminho_excel, dados_extraidos):
    """
    Atualiza a planilha Excel com os dados extraídos.
    """
    try:
        import openpyxl
    except ImportError:
        print("ERRO: Biblioteca openpyxl não instalada. Instale com: pip install openpyxl")
        return

    print(f"\nAtualizando planilha: {caminho_excel}...")
    
    try:
        # Carregar com keep_vba=True para não estragar as macros
        wb = openpyxl.load_workbook(caminho_excel, keep_vba=True)
    except Exception as e:
        print(f"ERRO ao abrir planilha: {e}")
        return

    alteracoes = 0
    
    for dados in dados_extraidos:
        tipo = dados['A_tipo_peca']
        titulo = dados['B_titulo_peca']
        
        # Mapear tipo para nome da aba
        if tipo == 'PILAR':
            sheet_name = 'Pilares'
        elif tipo == 'VIGA':
            sheet_name = 'Vigas'
        elif tipo == 'LAJE':
            sheet_name = 'Lajes'
        else:
            print(f"  Pular {titulo}: Tipo '{tipo}' não corresponde a uma aba conhecida.")
            continue
        
        if sheet_name not in wb.sheetnames:
            print(f"  ERRO: Aba '{sheet_name}' não existe na planilha.")
            continue
            
        ws = wb[sheet_name]
        
        # Procurar a linha pelo Título da Peça (Coluna B)
        # Assumindo que a coluna B é a 2ª coluna (índice 2 no openpyxl 1-based)
        linha_encontrada = None
        for row in range(4, ws.max_row + 1):  # Começa da linha 4 (dados)
            cell_val = ws.cell(row=row, column=2).value
            if cell_val and str(cell_val).strip() == titulo:
                linha_encontrada = row
                break
        
        if not linha_encontrada:
            print(f"  Peça '{titulo}' não encontrada. Criando nova linha...")
            # Encontrar próxima linha vazia a partir da linha 4
            # NÃO usar ws.max_row pois pode incluir linhas com formatação sem dados
            linha_encontrada = 4
            while ws.cell(row=linha_encontrada, column=2).value is not None:
                linha_encontrada += 1
            
            # Preencher Colunas de Identificação (A e B)
            ws.cell(row=linha_encontrada, column=1).value = tipo      # Coluna A: Tipo
            ws.cell(row=linha_encontrada, column=2).value = titulo    # Coluna B: Título
            
        # Colunas de dados brutos (extraídos do DXF)
        mapa_dados = {
            3: 'C_nome_desenho',         # C
            4: 'D_quantidade',           # D
            5: 'E_secao',               # E
            6: 'F_comprimento_cm',      # F
            7: 'G_volume_concreto_m3',  # G
            9: 'I_fck_mpa',             # I
            10: 'J_fcj_mpa',            # J
            11: 'K_peso_aco_frouxo_kg', # K
            12: 'L_peso_aco_protendido_kg', # L
            13: 'M_peso_aco_consolo_kg',    # M
        }

        r = linha_encontrada
        print(f"  Escrevendo '{titulo}' na linha {r}...")

        # Escrever dados brutos
        for col_idx, chave in mapa_dados.items():
            valor = dados.get(chave)
            if valor is not None:
                ws.cell(row=r, column=col_idx).value = valor

        # Escrever fórmulas (mesmo padrão das abas Vigas/Lajes)
        ws.cell(row=r, column=8).value = f'=G{r}*D{r}'              # H: Volume total
        ws.cell(row=r, column=14).value = f'=G{r}*2400+K{r}+L{r}+M{r}'  # N: Peso total unitário
        ws.cell(row=r, column=15).value = f'=(K{r}+M{r})/G{r}'     # O: Taxa aço frouxo
        ws.cell(row=r, column=16).value = f'=L{r}/G{r}'             # P: Taxa aço protendido
        ws.cell(row=r, column=17).value = f'=N{r}*D{r}/1000'        # Q: Peso total (t)
        ws.cell(row=r, column=18).value = f'=K{r}*D{r}/1000'        # R: Peso total frouxo (t)
        ws.cell(row=r, column=19).value = f'=L{r}*D{r}/1000'        # S: Peso total protendido (t)
        
        alteracoes += 1

    if alteracoes > 0:
        try:
            wb.save(caminho_excel)
            print(f"\nPlanilha salva com sucesso! {alteracoes} peças atualizadas.")
        except PermissionError:
            print(f"\nERRO DE PERMISSÃO: O arquivo '{caminho_excel}' está aberto.")
            print("POR FAVOR, FECHE O ARQUIVO EXCEL E TENTE NOVAMENTE.")
        except Exception as e:
            print(f"ERRO ao salvar planilha: {e}")
    else:
        print("\nNenhuma alteração feita na planilha.")


# =============================================================================
# MAIN
# =============================================================================

def main():
    if len(sys.argv) < 2:
        print("Uso: python extrair_dados_dxf.py <arquivo.dxf ou pasta> [caminho_planilha.xlsm]")
        print("  Extrai dados de desenhos DXF.")
        print("  Se o caminho da planilha for informado, ela será atualizada automaticamente.")
        sys.exit(1)

    caminho = sys.argv[1]
    
    # Verificar se foi passado o excel como segundo argumento
    caminho_excel = None
    if len(sys.argv) > 2:
        caminho_excel = sys.argv[2]
        if not os.path.exists(caminho_excel):
            print(f"AVISO: Planilha não encontrada: {caminho_excel}")
            caminho_excel = None

    if os.path.isdir(caminho):
        arquivos = sorted(Path(caminho).glob('*.dxf'))
        if not arquivos:
            print(f"Nenhum .dxf em: {caminho}")
            sys.exit(1)
        print(f"Encontrados {len(arquivos)} arquivo(s) DXF\n")
    else:
        arquivos = [Path(caminho)]

    todos = []
    for arq in arquivos:
        print(f"Processando: {arq.name}")
        dados = extrair_dados_completos(str(arq))
        if dados:
            print(formatar_relatorio(dados))
            todos.append(dados)
        else:
            print(f"  ERRO ao processar {arq.name}\n")

    # Salvar JSON
    if todos:
        base = Path(caminho).parent if os.path.isfile(caminho) else Path(caminho)
        out = base / 'dados_extraidos.json'
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(todos, f, ensure_ascii=False, indent=2)
        print(f"Dados salvos em: {out}")
        
        # Atualizar Excel se solicitado
        if caminho_excel:
            atualizar_planilha(caminho_excel, todos)

    return todos


if __name__ == '__main__':
    main()
