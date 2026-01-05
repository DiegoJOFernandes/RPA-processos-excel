# 📘 Manual de Uso — Automação de Faturas em Excel (Python)

## 1. Objetivo da automação

Esta automação tem como objetivo **gerar faturas automaticamente a partir de uma planilha Excel**, realizando:

* leitura dos dados de entrada;
* agrupamento por cliente (CPF/CNPJ);
* preenchimento de um template de fatura;
* geração de um arquivo de fatura por cliente;
* tentativa de impressão automática;
* organização dos arquivos em pastas.

---

## 2. Pré-requisitos

Antes de usar a automação, verifique se você possui:

* ✅ Windows
* ✅ Python instalado
* ✅ Ambiente virtual configurado (`.venv`)
* ✅ Dependências instaladas
* ✅ Microsoft Excel instalado (para impressão automática)

---

## 3. Estrutura de pastas esperada

A automação **só funciona corretamente** se a estrutura abaixo for respeitada:

```
invoice_excel_automation/
│
├── input/
│   └── dados.xlsx              ← planilha de dados de entrada
│
├── templates/
│   └── fatura.xlsx             ← template da fatura
│
├── output/
│   └── FATURA_<ID>/            ← pastas geradas automaticamente
│
├── src/
│   └── (arquivos do sistema)
│
├── .env
└── README.md
```

---

## 4. Como preparar a planilha de entrada (dados.xlsx)

### Aba esperada

* Nome da aba: **`Dados`** (ou conforme configurado no `.env`)

### Colunas obrigatórias

A planilha deve conter, no mínimo, as seguintes colunas:

| Coluna            | Descrição                    |
| ----------------- | ---------------------------- |
| documento_cliente | CPF ou CNPJ do cliente       |
| nome_cliente      | Nome ou Razão Social         |
| descricao         | Descrição do produto/serviço |
| quantidade        | Quantidade                   |
| valor_unitario    | Valor unitário               |

⚠️ **Importante**

* Os nomes das colunas não diferenciam maiúsculas/minúsculas
* Espaços extras são tratados automaticamente

---

## 5. Como preparar o template da fatura (fatura.xlsx)

O template deve conter:

* Uma aba chamada **`Fatura PJ`** (ou conforme `.env`)
* Células reservadas para:

  * documento
  * nome do cliente
  * data
  * total da fatura
* Uma área de tabela para itens (descrição, quantidade, valores)

⚠️ O layout pode ser personalizado, desde que as células configuradas no `.env` sejam respeitadas.

---

## 6. Configuração do arquivo `.env`

Antes de executar, revise o arquivo `.env`.

Exemplo básico:

```env
INPUT_FILE=./input/dados.xlsx
TEMPLATE_FILE=./templates/fatura.xlsx

SHEET_INPUT=Dados
SHEET_TEMPLATE=Fatura PJ

GROUP_BY_COLUMN=documento_cliente

ITEM_DESC_COLUMN=descricao
ITEM_QTY_COLUMN=quantidade
ITEM_UNIT_COLUMN=valor_unitario
ITEM_TOTAL_COLUMN=valor_total

OUTPUT_DIR=./output
```

⚠️ Se o nome do arquivo ou da aba for diferente, **ajuste aqui**.

---

## 7. Como executar a automação

### Passo 1 — Ativar o ambiente virtual

Na raiz do projeto:

```powershell
.\.venv\Scripts\Activate.ps1
```

Você saberá que deu certo quando aparecer:

```
(.venv)
```

---

### Passo 2 — Executar a automação

```powershell
python -m src.main ou .\.venv\Scripts\python.exe -m src.main
```

---

## 8. O que acontece durante a execução

1. O sistema verifica se o arquivo de entrada existe
2. Lê e valida os dados
3. Agrupa os registros por cliente
4. Para cada cliente:

   * cria uma pasta em `output/`
   * gera uma fatura em Excel
   * tenta imprimir
   * salva um arquivo `status.txt` com o resultado

---

## 9. Estrutura do resultado (output)

Exemplo:

```
output/
└── FATURA_12345678000199/
    ├── fatura_12345678000199.xlsx
    └── status.txt
```

Conteúdo do `status.txt`:

* `PRINT_OK` → impressão realizada com sucesso
* `PRINT_FAIL: <motivo>` → impressão falhou, mas arquivo foi salvo

---

## 10. Mensagens de erro comuns e como resolver

### ❌ Arquivo de entrada não encontrado

**Causa:** nome errado ou arquivo fora da pasta `input/`

**Solução:**

* Verifique o nome do arquivo
* Ajuste o `.env`

---

### ❌ Template de fatura não encontrado

**Causa:** arquivo inexistente ou nome incorreto

**Solução:**

* Verifique a pasta `templates/`
* Ajuste `TEMPLATE_FILE` no `.env`

---

### ❌ Aba não encontrada

**Causa:** nome da aba diferente do configurado

**Solução:**

* Abra o Excel
* Copie exatamente o nome da aba
* Atualize o `.env`

---

## 11. Boas práticas de uso

* ✔️ Sempre feche o Excel antes de rodar a automação
* ✔️ Não altere a estrutura das colunas sem avisar
* ✔️ Execute uma vez e valide antes de rodar em lote grande
* ✔️ Guarde o `output/` como evidência

---

## 12. Observações finais

* A automação **não altera o arquivo de entrada**
* Cada execução gera novas faturas
* Erros são tratados de forma controlada e exibidos no terminal

---

## 13. Suporte / Evoluções futuras

Possíveis melhorias:

* suporte a PF e PJ
* exportação automática para PDF
* execução agendada
* integração com sistemas externos

