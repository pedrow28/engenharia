# Tech Estrutural · Automação de Projetos

Sistema que lê automaticamente os dados técnicos dos desenhos de engenharia (DWG) e preenche a planilha de controle — sem precisar abrir nenhum arquivo manualmente.

---

## O que o sistema faz

Para cada desenho de **Pilar**, **Viga** ou **Laje**, o sistema extrai:

| Dado | Exemplo |
|---|---|
| Tipo e título da peça | Pilar P5, Viga VR720, Laje L-1001 |
| Quantidade e seção | 2 peças, 30×80 cm |
| Comprimento | 1408 cm |
| Volume de concreto | 3,66 m³ |
| fck e fcj | 45 MPa / 31,5 MPa |
| Peso do aço frouxo | 378 kg |
| Peso do aço protendido | 33 kg |
| Taxas e pesos totais | calculados automaticamente |

Tudo isso vai direto para a planilha `.xlsm` — sem digitação manual.

---

## Pré-requisitos

Você precisa instalar **dois programas** antes de usar:

### 1. Python

> Se já tiver o Python instalado, pule esta etapa.

1. Acesse: **python.org/downloads**
2. Clique em **"Download Python 3.x"**
3. Execute o instalador
4. **IMPORTANTE:** marque a opção **"Add Python to PATH"** antes de clicar em *Install Now*

Para verificar: abra o Prompt de Comando e digite `python --version`. Deve aparecer algo como `Python 3.12.x`.

---

### 2. ODA File Converter

Programa gratuito que converte arquivos DWG para o formato que o sistema lê.

1. Acesse: **opendesign.com/guestfiles/oda_file_converter**
2. Baixe a versão para **Windows**
3. Instale no caminho padrão (não altere)

---

## Instalação do sistema

Faça isso **uma única vez** após baixar o projeto.

1. Abra o **Prompt de Comando** na pasta do projeto
   - Navegue até a pasta onde o projeto está salvo
   - Clique na barra de endereço do Windows Explorer, digite `cmd` e pressione Enter

2. Execute o comando abaixo e aguarde terminar:

```
pip install ezdxf openpyxl fastapi uvicorn
```

Pronto. O sistema está pronto para uso.

---

## Como usar

### Iniciando a interface

Na pasta do projeto, dê **duplo clique** no arquivo `iniciar.bat`

— ou —

Abra o Prompt de Comando na pasta do projeto e digite:

```
python app.py
```

O navegador abrirá automaticamente em `http://localhost:8000` com a interface do sistema.

---

### Passo a passo na interface

**1. Selecionar as pastas dos desenhos**

Clique em **"…"** ao lado de cada tipo de peça (Pilares, Vigas, Lajes) e escolha a pasta que contém os arquivos DWG correspondentes. Você pode preencher apenas os tipos que precisa processar — não é obrigatório preencher todos.

**2. Selecionar a planilha**

Clique em **"Selecionar"** ao lado de *Planilha de Controle* e escolha o arquivo `.xlsm`.

> ⚠️ **A planilha deve estar fechada no Excel** antes de executar. Se estiver aberta, o sistema não conseguirá salvar e dará erro de permissão.

**3. Executar**

Clique no botão verde **"Executar Automação"**.

O log na parte inferior da tela mostrará o progresso em tempo real. Ao final, aparecerá a mensagem **"AUTOMAÇÃO CONCLUÍDA COM SUCESSO"**.

---

## Fluxo do processamento

```
Pasta DWG selecionada
        │
        ▼
 Conversão DWG → DXF
  (ODA File Converter)
        │
        ▼
  Extração de dados
   (leitura do DXF)
        │
        ▼
  Planilha .xlsm
   atualizada
```

O sistema converte e lê cada arquivo automaticamente — você não precisa fazer nada além de clicar em Executar.

---

## Solução de problemas

**"ODA File Converter não encontrado"**
O ODA File Converter não está instalado ou foi instalado em um caminho diferente do padrão. Reinstale sem alterar o diretório de destino.

**"Planilha não encontrada" ou erro de permissão**
Feche o arquivo Excel completamente antes de executar o sistema.

**Valores zerados ou campos em branco na planilha**
Os nomes dos blocos dentro do AutoCAD precisam estar no padrão da empresa (`NOTAS`, `CARIMBO`, `SM_formatoA4paraLajes`, etc.). Verifique com quem gerou o desenho.

**O navegador não abre automaticamente**
Abra manualmente e acesse: `http://localhost:8000`

**Erro ao instalar as dependências**
Certifique-se de que o Python foi instalado com a opção "Add to PATH" marcada. Se necessário, desinstale e reinstale o Python com essa opção ativada.

---

## Estrutura do projeto

```
engenharia/
├── app.py                  ← Servidor da interface web
├── templates/
│   └── index.html          ← Interface visual
├── scripts/
│   ├── extrair_dados_dxf.py   ← Motor de extração
│   ├── converter_dwg_dxf.py   ← Conversão DWG → DXF
│   └── requirements.txt
├── assets/
│   └── logo.png
└── dados/
    └── CONTROLE DE PEÇAS - CROMA.xlsm
```

---

*Tech Estrutural Projetos*
