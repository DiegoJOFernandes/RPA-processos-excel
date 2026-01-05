import pandas as pd
from datetime import date, datetime
from src.config import settings

REQUIRED_COLS = [
    settings.group_by_column.lower(),
    settings.item_desc_column.lower(),
    settings.item_qty_column.lower(),
    settings.item_unit_column.lower(),
]


def validate_and_clean(df: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes: {missing}")

    df = df.copy()

    # Remove linhas totalmente vazias
    df = df.dropna(how="all")

    # Normaliza campos essenciais
    key = settings.group_by_column.lower()
    df[key] = df[key].fillna("").astype(str).str.strip()

    # Remove linhas sem documento
    df = df[df[key] != ""]

    # Quantidade e valor (tentativa de normalizar para número)
    qty_col = settings.item_qty_column.lower()
    unit_col = settings.item_unit_column.lower()

    df[qty_col] = df[qty_col].fillna("0").astype(str).str.replace(",", ".")
    df[unit_col] = df[unit_col].fillna("0").astype(str).str.replace(",", ".")

    # Converte para float (onde possível)
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    df[unit_col] = pd.to_numeric(df[unit_col], errors="coerce").fillna(0)

    # Valor total por linha
    df[settings.item_total_column.lower()] = df[qty_col] * df[unit_col]

    return df


def group_invoices(df: pd.DataFrame):
    key = settings.group_by_column.lower()
    for doc, group in df.groupby(key, dropna=False):
        yield doc, group.reset_index(drop=True)


def invoice_header_from_group(doc: str, group: pd.DataFrame) -> dict:
    """
    Monta o cabeçalho da fatura baseado no grupo (cliente).
    Inclui campos extras: mês referência, número do cartão e total mensal.
    """

    # Nome do cliente (se existir)
    nome = ""
    if "nome_cliente" in group.columns:
        nome = str(group["nome_cliente"].iloc[0]).strip()

    # Mês de referência e número do cartão (com base nas colunas configuradas)
    month_ref_col = settings.month_ref_column.lower()
    card_col = settings.card_number_column.lower()
    monthly_sum_col = settings.monthly_sum_column.lower()

    month_ref = str(group[month_ref_col].iloc[0]).strip() if month_ref_col in group.columns else ""
    card_number = str(group[card_col].iloc[0]).strip() if card_col in group.columns else ""

    # Total mensal: se a coluna já vem preenchida, usa ela; senão soma o valor_total
    if monthly_sum_col in group.columns:
        raw = str(group[monthly_sum_col].iloc[0]).strip().replace(",", ".")
        try:
            total_mensal = float(raw) if raw != "" else 0.0
        except ValueError:
            total_mensal = 0.0
    else:
        # fallback: soma o valor_total (já calculado no validate_and_clean)
        total_mensal = float(group[settings.item_total_column.lower()].astype(float).sum())

    total_mensal = round(total_mensal, 2)

    return {
        "documento": doc,
        "nome": nome,
        "data_emissao": datetime.now().strftime("%d/%m/%Y"),

        # total da fatura (usando total mensal)
        "total": total_mensal,

        # novos campos
        "mes_referencia": month_ref,
        "numero_cartao": card_number,
        "total_mensal": total_mensal,
    }
