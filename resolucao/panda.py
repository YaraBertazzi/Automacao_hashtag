import pandas as pd  #pd é um apaledido para pandas

tabela = pd.read_excel(r'C:\Users\HD\Downloads/Vendas - Dez.xlsx') #endereço da planilha do exel para importar
print(tabela)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
print(f'{quantidade}, e {faturamento}')