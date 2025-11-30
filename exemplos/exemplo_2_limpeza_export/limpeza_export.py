from pathlib import Path
import pandas as pd


def ler_csv_bruto(caminho: Path) -> pd.DataFrame:
    """
    Lê o CSV bruto tentando lidar com separador ;/, e encodings comuns (utf-8 / latin-1 / cp1252).
    Devolve um DataFrame com as colunas tal como vêm no ficheiro.
    """
    for sep in [";", ","]:
        for encoding in ["utf-8", "latin-1", "cp1252"]:
            try:
                df = pd.read_csv(caminho, sep=sep, encoding=encoding)
                # Se só vier 1 coluna, provavelmente o separador está errado
                if df.shape[1] > 1:
                    print(
                        f"DEBUG: lido com sep='{sep}', encoding='{encoding}', colunas={list(df.columns)}")
                    return df
            except Exception:
                continue

    # Último recurso
    df = pd.read_csv(caminho, encoding="utf-8", engine="python")
    print(f"DEBUG: fallback, colunas={list(df.columns)}")
    return df


def limpar_dados(df: pd.DataFrame) -> pd.DataFrame:
    """Limpa e normaliza os dados do export bruto."""

    # Guardar nomes originais para debug
    colunas_originais = list(df.columns)

    # Tira espaços aos nomes das colunas
    df.columns = [c.strip() for c in df.columns]

    # Mapeamento de nomes "porcos" para nomes limpos
    mapeamento = {
        "data": "Data",
        "cliente": "Cliente",
        "produto": "Produto",
        "quantidade": "Quantidade",
        "qtd": "Quantidade",
        "total": "Total",
        "valor": "Total",
    }

    colunas_normalizadas = {}
    for c in df.columns:
        chave = c.strip().lower()
        if chave in mapeamento:
            colunas_normalizadas[c] = mapeamento[chave]

    # Renomeia colunas
    df = df.rename(columns=colunas_normalizadas)

    colunas_desejadas = ["Data", "Cliente", "Produto", "Quantidade", "Total"]
    colunas_existentes = [c for c in colunas_desejadas if c in df.columns]

    if len(colunas_existentes) < len(colunas_desejadas):
        em_falta = set(colunas_desejadas) - set(colunas_existentes)
        raise ValueError(
            f"Faltam colunas obrigatórias no ficheiro.\n"
            f"Originais: {colunas_originais}\n"
            f"Após normalização: {list(df.columns)}\n"
            f"Em falta: {em_falta}"
        )

    # Mantém só as colunas que nos interessam
    df = df[colunas_desejadas].copy()

    # Remove linhas completamente vazias
    df = df.dropna(how="all")

    # Strip de espaços em texto
    for col in ["Cliente", "Produto"]:
        df[col] = df[col].astype(str).str.strip()

    # Normaliza cliente/produto (primeira letra maiúscula)
    df["Cliente"] = df["Cliente"].str.title()
    df["Produto"] = df["Produto"].str.title()

    # Converte Data para datetime (assumindo formato dia-mês-ano)
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")

    # Troca vírgula por ponto em campos numéricos e converte
    for col in ["Quantidade", "Total"]:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Remove linhas sem cliente ou sem Total válido
    df = df.dropna(subset=["Cliente", "Total"])

    # Remove duplicados exactos (se houver)
    df = df.drop_duplicates()

    return df


def exportar_excel(df_limpo: pd.DataFrame, pasta_saida: Path) -> Path:
    """Exporta os dados limpos para um ficheiro Excel."""
    caminho_saida = pasta_saida / "dados_limpos.xlsx"
    df_limpo.to_excel(caminho_saida, index=False)
    return caminho_saida


def main() -> None:
    base_dir = Path(__file__).resolve().parent
    caminho_csv = base_dir / "export_bruto.csv"

    if not caminho_csv.exists():
        raise FileNotFoundError(f"Ficheiro não encontrado: {caminho_csv}")

    print(f"Lendo export bruto: {caminho_csv}")
    df_bruto = ler_csv_bruto(caminho_csv)

    print(f"Linhas brutas: {len(df_bruto)}")
    df_limpo = limpar_dados(df_bruto)
    print(f"Linhas após limpeza: {len(df_limpo)}")

    caminho_saida = exportar_excel(df_limpo, base_dir)
    print(f"Ficheiro limpo criado: {caminho_saida}")


if __name__ == "__main__":
    main()
