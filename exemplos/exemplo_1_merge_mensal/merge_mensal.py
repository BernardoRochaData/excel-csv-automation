from pathlib import Path
import pandas as pd


def carregar_dados(pasta_dados: Path) -> pd.DataFrame:
    """Lê todos os ficheiros .xlsx da pasta e junta num único DataFrame."""
    ficheiros = list(pasta_dados.glob("*.xlsx"))

    print(f"Pasta de dados: {pasta_dados}")
    print(f"Ficheiros encontrados: {[f.name for f in ficheiros]}")

    if not ficheiros:
        raise FileNotFoundError(f"Nenhum .xlsx encontrado em {pasta_dados}")

    lista_df = []
    for ficheiro in ficheiros:
        print(f"Lendo ficheiro: {ficheiro.name}")
        df = pd.read_excel(ficheiro)

        colunas_obrigatorias = {"Data", "Cliente",
                                "Produto", "Quantidade", "Total"}
        if not colunas_obrigatorias.issubset(df.columns):
            raise ValueError(
                f"Ficheiro {ficheiro.name} não tem as colunas necessárias: "
                f"{colunas_obrigatorias}"
            )

        lista_df.append(df)

    df_todos = pd.concat(lista_df, ignore_index=True)
    df_todos["Data"] = pd.to_datetime(df_todos["Data"])

    return df_todos


def gerar_relatorios(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)

    resumo_clientes = (
        df.groupby("Cliente", as_index=False)["Total"]
        .sum()
        .sort_values("Total", ascending=False)
    )

    resumo_produtos = (
        df.groupby("Produto", as_index=False)["Total"]
        .sum()
        .sort_values("Total", ascending=False)
    )

    resumo_mensal = (
        df.groupby("AnoMes", as_index=False)["Total"]
        .sum()
        .sort_values("AnoMes")
    )

    return {
        "Dados_Limpos": df,
        "Resumo_Clientes": resumo_clientes,
        "Resumo_Produtos": resumo_produtos,
        "Resumo_Mensal": resumo_mensal,
    }


def exportar_para_excel(relatorios: dict[str, pd.DataFrame], nome_ficheiro: str) -> None:
    with pd.ExcelWriter(nome_ficheiro, engine="openpyxl") as writer:
        for nome_sheet, df in relatorios.items():
            df.to_excel(writer, sheet_name=nome_sheet, index=False)

    print(f"Relatório criado com sucesso: {nome_ficheiro}")


def main() -> None:
    # Pasta onde está o ficheiro merge_mensal.py
    base_dir = Path(__file__).resolve().parent
    pasta_dados = base_dir / "dados_brutos"

    df_todos = carregar_dados(pasta_dados)
    relatorios = gerar_relatorios(df_todos)
    exportar_para_excel(relatorios, base_dir / "relatorio_anual.xlsx")


if __name__ == "__main__":
    main()
