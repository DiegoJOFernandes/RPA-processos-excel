import pandas as pd
from pathlib import Path
from src.config import settings


def read_input_excel() -> pd.DataFrame:
    """
    Lê o arquivo Excel de entrada e retorna um DataFrame padronizado.

    Raises:
        FileNotFoundError: se o arquivo de entrada não existir
    """

    # Converte o caminho configurado em um objeto Path
    input_path = Path(settings.input_file)

    # Verifica se o arquivo realmente existe antes de tentar ler
    if not input_path.exists():
        print("❌ ERRO: Arquivo de entrada não encontrado.")
        print(f"   Caminho esperado: {input_path.resolve()}")
        print("   Verifique se o arquivo existe e se o nome está correto no .env")
        raise FileNotFoundError(f"Arquivo não encontrado: {input_path}")

    print(f"📂 Lendo arquivo de entrada: {input_path.resolve()}")

    # Lê o Excel com pandas
    df = pd.read_excel(
        input_path,
        sheet_name=settings.sheet_input,
        dtype=str
    )

    # Normaliza os nomes das colunas
    df.columns = [c.strip().lower() for c in df.columns]

    print(f"✅ Arquivo lido com sucesso ({len(df)} linhas).")

    return df
