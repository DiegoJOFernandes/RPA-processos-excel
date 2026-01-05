```markdown
# 📄 Invoice Excel Automation (Python)

Projeto em Python para **automatizar a geração de faturas em Excel**, a partir de uma planilha de dados de entrada.  
O processo lê os dados, agrupa por cliente, preenche um template de fatura, salva o arquivo final em pasta específica e realiza a impressão automática (Windows).

---

## 🎯 Objetivo

Automatizar o processo manual de:
- leitura de planilhas Excel,
- agrupamento de dados por cliente,
- preenchimento de um template de fatura,
- geração de arquivos finais,
- impressão das faturas,
- organização dos arquivos gerados.

---

## ⚙️ Funcionalidades

- Leitura de planilha Excel de entrada
- Validação e tratamento de dados
- Agrupamento por cliente (CPF/CNPJ ou outro identificador)
- Preenchimento automático de template de fatura em Excel
- Cálculo de totais
- Geração de uma fatura por cliente
- Impressão automática da fatura (Windows + Excel instalado)
- Salvamento organizado em pastas por fatura

---

## 🧱 Estrutura do Projeto

```

invoice_excel_automation/
│
├── input/
│   └── dados.xlsx              # Planilha de entrada
│
├── templates/
│   └── fatura.xlsx             # Template da fatura
│
├── output/
│   └── FATURA_<ID>/            # Faturas geradas
│       └── fatura_<ID>.xlsx
│
├── src/
│   ├── config.py               # Configurações do projeto
│   ├── io_excel.py             # Leitura do Excel de entrada
│   ├── transform.py            # Validação, limpeza e agrupamento
│   ├── fill_template.py        # Preenchimento do template
│   ├── print_invoice.py        # Impressão da fatura
│   └── main.py                 # Orquestração do processo
│
├── requirements.txt            # Dependências
├── .env                        # Configurações de ambiente
└── README.md

````

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.12+**
- **pandas** – leitura e manipulação de dados
- **openpyxl** – leitura e escrita em Excel
- **python-dotenv** – variáveis de ambiente
- **pywin32** – impressão automática (somente Windows)

---

## 📦 Instalação

### 1. Criar ambiente virtual
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
````

### 2. Instalar dependências

```powershell
pip install -r requirements.txt
```

---

## 🔧 Configuração

Edite o arquivo `.env` conforme o layout das suas planilhas:

```env
INPUT_FILE=./input/dados.xlsx
TEMPLATE_FILE=./templates/fatura.xlsx
OUTPUT_DIR=./output

SHEET_INPUT=Dados
SHEET_TEMPLATE=Fatura PJ

GROUP_BY_COLUMN=documento_cliente
ITEM_DESC_COLUMN=descricao
ITEM_QTY_COLUMN=quantidade
ITEM_UNIT_COLUMN=valor_unitario
ITEM_TOTAL_COLUMN=valor_total

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
```

---

## ▶️ Execução

Com o ambiente virtual ativado:

```powershell
python src/main.py
```

---

## 📄 Resultado Esperado

Após a execução:

* Será criada uma pasta para cada cliente em `output/`
* Cada pasta conterá a fatura preenchida em Excel
* A fatura será enviada para impressão automaticamente (se disponível)
* Um arquivo `status.txt` indica sucesso ou falha de impressão

Exemplo:

```
output/
└── FATURA_12345678000199/
    ├── fatura_12345678000199.xlsx
    └── status.txt
```

---

## 🖨️ Impressão Automática

* Disponível apenas no **Windows**
* Requer **Microsoft Excel instalado**
* Usa a impressora padrão do sistema

Caso a impressão falhe, o arquivo da fatura permanece salvo para impressão manual.

---

## ⚠️ Observações Importantes

* Linhas inválidas são removidas automaticamente
* Se um cliente não possuir linhas válidas, a fatura não é gerada
* O projeto foi pensado para **uso operacional simples**, sem banco de dados ou APIs

---

## 🚀 Próximas Evoluções (opcional)

* Suporte a PF e PJ com templates diferentes
* Exportação automática para PDF
* Geração de executável (.exe)
* Logs estruturados
* Integração com sistemas externos
* Agendamento automático (Task Scheduler)

---

## 👤 Autor / Responsável

Projeto desenvolvido para automação de processos internos com Excel utilizando Python.

---
