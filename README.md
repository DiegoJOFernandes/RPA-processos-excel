```md
# 🤖 RPA – Geração de Faturas de Cartão de Crédito (PF / PJ)

Este projeto é uma automação (**RPA**) desenvolvida em **Python** para gerar **faturas de cartão de crédito** a partir de uma planilha de transações, com suporte a:

- Pessoa Física (**PF**)  
- Pessoa Jurídica (**PJ**)  
- Templates distintos de fatura  
- Geração de **Excel + PDF**  
- Organização automática de arquivos  
- Validações completas antes da execução (preflight)

O projeto foi pensado para uso **corporativo**, com foco em confiabilidade, rastreabilidade e fácil manutenção.

---

## 🎯 Objetivo

Automatizar o processo de:

1. Leitura de uma planilha de transações de cartão de crédito
2. Agrupamento por cliente (CPF ou CNPJ)
3. Identificação automática de PF ou PJ
4. Cálculo do total mensal
5. Preenchimento de templates de fatura em Excel
6. Geração do PDF da fatura
7. Organização dos arquivos por cliente
8. Execução segura com validações prévias

---

## 🧱 Arquitetura do Projeto

```

invoice_excel_automation/
│
├── src/
│   ├── main.py                 # Orquestra o fluxo principal do RPA
│   ├── config.py               # Configurações centralizadas (via .env)
│   ├── io_excel.py             # Leitura da planilha de entrada
│   ├── transform.py            # Validações, agrupamentos e header da fatura
│   ├── fill_template.py        # Preenchimento do template Excel (PF/PJ)
│   ├── print_invoice.py        # Exportação para PDF e impressão (Windows)
│   └── preflight.py            # Validações antes de iniciar o RPA
│
├── input/                      # Planilha de dados (não versionar)
├── templates/                  # Templates de fatura PF e PJ
├── output/                     # Faturas geradas automaticamente
│
├── .env                        # Configurações de ambiente
├── .gitignore
├── requirements.txt
└── README.md

````

---

## ⚙️ Pré-requisitos

- **Python 3.10+**
- **Windows** (para exportação PDF via Excel)
- Microsoft **Excel instalado** (para PDF/print)
- Git (opcional)

---

## 📦 Instalação

### 1️⃣ Clonar o repositório
```bash
git clone <url-do-repositorio>
cd invoice_excel_automation
````

### 2️⃣ Criar ambiente virtual

```bash
python -m venv .venv
```

### 3️⃣ Ativar ambiente virtual

**Windows (PowerShell):**

```powershell
.\.venv\Scripts\Activate.ps1
```

### 4️⃣ Instalar dependências

```bash
pip install -r requirements.txt
```

---

## 🔐 Configuração (`.env`)

Crie um arquivo `.env` na raiz do projeto com o seguinte conteúdo:

```env
INPUT_FILE=./input/dados.xlsx
TEMPLATE_PF=./templates/fatura_pf.xlsx
TEMPLATE_PJ=./templates/fatura_pj.xlsx
OUTPUT_DIR=./output

SHEET_INPUT=Dados
SHEET_TEMPLATE=Fatura

CLIENT_TYPE_COLUMN=tipo_cliente
GROUP_BY_COLUMN=documento_cliente

MONTH_REF_COLUMN=mes_fatura
CARD_NUMBER_COLUMN=numero_cartao
MONTHLY_SUM_COLUMN=soma_total_mensal

MAX_ITEMS=40

CELL_DOC=B6
CELL_NAME=B7
CELL_DATE=B8
CELL_TOTAL=H25

ITEMS_START_ROW=12
COL_ITEM_DESC=B
COL_ITEM_QTY=F
COL_ITEM_UNIT=G
COL_ITEM_TOTAL=H

CELL_MONTH_REF=D6
CELL_CARD_NUMBER=D7
CELL_MONTHLY_SUM=D8
```

---

## 📥 Planilha de Entrada (Input)

A planilha deve conter **uma linha por transação** com as colunas abaixo:

### 🔑 Colunas obrigatórias

| Coluna            | Descrição                       |
| ----------------- | ------------------------------- |
| documento_cliente | CPF ou CNPJ                     |
| tipo_cliente      | `PF` ou `PJ`                    |
| nome_cliente      | Nome do cliente                 |
| mes_fatura        | Mês de referência (ex: 08/2024) |
| numero_cartao     | Número do cartão                |
| estabelecimento   | Nome do estabelecimento         |
| valor_compra      | Valor total da compra           |
| qtd_parcelas      | Quantidade de parcelas          |
| valor_parcela     | Valor da parcela mensal         |

---

## 🧾 Templates de Fatura

* `templates/fatura_pf.xlsx`
* `templates/fatura_pj.xlsx`

### Requisitos:

* Devem conter a aba **`Fatura`**
* Podem conter **células mescladas**
* Células devem respeitar as posições configuradas no `.env`

O sistema trata automaticamente células mescladas.

---

## ✅ Preflight Checks (Validações Iniciais)

Antes de qualquer processamento, o sistema valida:

* Existência do arquivo de input
* Existência dos templates PF e PJ
* Aba correta no template
* Colunas obrigatórias
* Valores válidos (`PF` / `PJ`)
* Documento preenchido
* Valores numéricos coerentes
* Quantidade total de faturas a gerar

Se algo estiver errado, o processo **é interrompido imediatamente** com erro claro.

---

## ▶️ Execução do RPA

Com tudo configurado, execute:

```bash
python -m src.main
```

---

## 📤 Estrutura de Saída

O sistema gera a seguinte estrutura automaticamente:

```
output/
└── PF/
    └── FATURA_12345678900/
        ├── fatura_12345678900.xlsx
        ├── fatura_12345678900.pdf
        └── status.txt
```

Ou:

```
output/
└── PJ/
    └── FATURA_12345678000199/
```

---

## 🖨️ PDF e Impressão

* A exportação para **PDF A4** é feita via Excel (Windows)
* Impressão automática é opcional
* Em outros sistemas operacionais, o PDF pode ser gerado futuramente via LibreOffice

---

## 🛡️ Boas Práticas Aplicadas

* Fail fast (erros antes do processamento)
* Configuração centralizada
* Templates desacoplados do código
* Código defensivo (merged cells, arquivos ausentes)
* Organização clara de saída
* Estrutura pronta para escalar

---

## 🚀 Evoluções Futuras (opcional)

* Modo `--dry-run`
* Logs estruturados
* Executável (`pyinstaller`)
* Validação CPF/CNPJ
* Integração com sistemas web
* Agendamento automático
* Interface gráfica (RPA visual)

---

## 📄 Licença

Projeto interno / uso corporativo.

---