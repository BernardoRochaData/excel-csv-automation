import pandas as pd

# Dados de exemplo
dados = {
    "Cliente": ["Cliente A", "Cliente B", "Cliente A", "Cliente C"],
    "Total": [100, 200, 50, 300],
}

# Criar DataFrame
df = pd.DataFrame(dados)

# Criar relatório: total por cliente
relatorio = df.groupby("Cliente")["Total"].sum().reset_index()

# Exportar para Excel
relatorio.to_excel("relatorio_clientes.xlsx", index=False)

print("Relatório criado com sucesso: relatorio_clientes.xlsx")
