import pandas as pd
import win32com.client as win32

#importa a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max.columns', None)
print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby("ID Loja").sum()
print(quantidade)

print('-' * 50)
# ticket médio por produtos em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'araodeveloper@gmail.com'
mail.Subject = 'Olha essa merda ai corno'
mail.HTMLBody = '''
Prezado,

Relatório de vendas

Faturamentos:
{}

Quantidade de vendas:
{}

ticket Médio dso Produtos em cada Loja:
{}

'''

mail.Send()

print("Email enviado!")
