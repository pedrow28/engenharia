# Automação de Extração de Dados: AutoCAD → Excel

Este projeto automatiza a extração de dados técnicos (fck, pesos de aço, volume, etc.) dos desenhos DWG para preenchimento da planilha de controle.

## Pré-requisitos

1.  **Python 3.x** instalado.
2.  **ODA File Converter** instalado (Gratuito).
    -   Download: [opendesign.com/guestfiles/oda_file_converter](https://www.opendesign.com/guestfiles/oda_file_converter)
3.  Biblioteca `ezdxf` instalada:
    ```bash
    pip install ezdxf
    ```

## Como Usar

### Passo 1: Converter Desenhos (DWG → DXF)

O Python precisa do formato DXF para ler os dados. Use este script para converter uma pasta inteira ou um arquivo:

```bash
# Converter uma pasta inteira (cria arquivos .dxf na mesma pasta)
python converter_dwg_dxf.py "C:\Caminho\Para\Os\Desenhos"

# Ou converter um arquivo específico
python converter_dwg_dxf.py "C:\Caminho\Desenho.dwg"
```

### Opção: Usar Interface Gráfica

Para facilitar, você pode usar a interface visual:

```bash
python interface_automacao.py
```

Uma janela abrirá onde você poderá selecionar a pasta e a planilha com o mouse e clicar em **EXECUTAR AUTOMAÇÃO**.

### Opção: Linha de Comando

Se preferir usar o terminal:

**1. Apenas Visualizar Dados:**
```bash
python extrair_dados_dxf.py "C:\Caminho\Para\Os\Desenhos"
```

**2. Atualizar Planilha Automaticamente:**
```bash
python extrair_dados_dxf.py "C:\Caminho\Para\Os\Desenhos" "C:\Caminho\Planilha.xlsm"
```

> **Importante**: A planilha **DEVE ESTAR FECHADA** para que a atualização funcione. Se estiver aberta, dará erro de permissão.

## Funcionalidades

O sistema extrai automaticamente:

*   **Identificação**: Tipo de peça (Pilar/Viga/Laje), Título (ex: P5), Seção, Comprimento.
*   **Concreto**: fck, Volume, Peso Próprio.
*   **Aço**:
    *   **Aço Frouxo**: Peso total da armação principal.
    *   **Aço Consolo**: Identifica automaticamente tabelas de consolo e aplica fator de simetria (2x) quando necessário.
*   **Cálculos Automáticos**: Taxas de aço (kg/m³) e Peso Total (t).

## Solução de Problemas

*   **Erro "ODA File Converter não encontrado"**: Certifique-se de que instalou o programa no caminho padrão (`C:\Program Files\ODA\...`).
*   **Valores zerados**: Verifique se os blocos dentro do AutoCAD estão com os nomes padrão (`NOTAS`, `CARIMBO`, `LISTA DE FERROS CONSOLO`). O script depende desses nomes.
