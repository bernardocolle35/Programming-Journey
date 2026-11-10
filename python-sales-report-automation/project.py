import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel("vendas.xlsx")
pd.set_option("display.max_columns", None)

faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
qnt_produtos = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
ticket_medio = (faturamento["Valor Final"] / qnt_produtos["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = '***************@gmail.com'
mail.subject = 'Relatorio de Venda por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>
<p>Segue o relatório de vendas por cada loja:</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{qnt_produtos.to_html()}

<p>Ticket Medio Dos Produtos Em Cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer duvida, estou a disposiçao</p>
<p>Atenciosamente: ********</p>
'''

mail.send()