from dataclasses import dataclass
import os
from dotenv import load_dotenv

# Carrega variáveis do arquivo .env
load_dotenv()


@dataclass(frozen=True)
class Settings:
    # ===============================
    # Arquivos e diretórios
    # ===============================
    input_file: str = os.getenv("INPUT_FILE", "./input/dados.xlsx")

    # Templates separados para PF e PJ
    template_pf: str = os.getenv("TEMPLATE_PF", "./templates/fatura_pf.xlsx")
    template_pj: str = os.getenv("TEMPLATE_PJ", "./templates/fatura_pj.xlsx")

    output_dir: str = os.getenv("OUTPUT_DIR", "./output")

    # ===============================
    # Planilhas / abas
    # ===============================
    sheet_input: str = os.getenv("SHEET_INPUT", "Dados")
    sheet_template: str = os.getenv("SHEET_TEMPLATE", "Fatura")

    # ===============================
    # Colunas de controle
    # ===============================
    group_by_column: str = os.getenv("GROUP_BY_COLUMN", "documento_cliente")

    # Define se o cliente é PF ou PJ
    client_type_column: str = os.getenv("CLIENT_TYPE_COLUMN", "tipo_cliente")

    # ===============================
    # Colunas de itens
    # ===============================
    item_desc_column: str = os.getenv("ITEM_DESC_COLUMN", "descricao")
    item_qty_column: str = os.getenv("ITEM_QTY_COLUMN", "quantidade")
    item_unit_column: str = os.getenv("ITEM_UNIT_COLUMN", "valor_unitario")
    item_total_column: str = os.getenv("ITEM_TOTAL_COLUMN", "valor_total")

    max_items: int = int(os.getenv("MAX_ITEMS", "40"))

    # ===============================
    # Células do template (comum PF/PJ)
    # ===============================
    cell_doc: str = os.getenv("CELL_DOC", "B6")
    cell_name: str = os.getenv("CELL_NAME", "B7")
    cell_date: str = os.getenv("CELL_DATE", "B8")
    cell_total: str = os.getenv("CELL_TOTAL", "H25")

    # ===============================
    # Tabela de itens no template
    # ===============================
    items_start_row: int = int(os.getenv("ITEMS_START_ROW", "12"))
    col_item_desc: str = os.getenv("COL_ITEM_DESC", "B")
    col_item_qty: str = os.getenv("COL_ITEM_QTY", "F")
    col_item_unit: str = os.getenv("COL_ITEM_UNIT", "G")
    col_item_total: str = os.getenv("COL_ITEM_TOTAL", "H")

    # ===============================
    # Campos extras da fatura (cartão / mês)
    # ===============================
    # Atenção: estas células precisam existir no template.
    cell_month_ref: str = os.getenv("CELL_MONTH_REF", "D6")
    cell_card_number: str = os.getenv("CELL_CARD_NUMBER", "D7")
    cell_monthly_sum: str = os.getenv("CELL_MONTHLY_SUM", "D8")

    # Colunas no Excel de entrada
    month_ref_column: str = os.getenv("MONTH_REF_COLUMN", "mes_fatura")
    card_number_column: str = os.getenv("CARD_NUMBER_COLUMN", "numero_cartao")
    monthly_sum_column: str = os.getenv("MONTHLY_SUM_COLUMN", "soma_total_mensal")


# Instância única de configuração
settings = Settings()
