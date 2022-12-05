import pandas as pd #importou com o apelido de pd
import win32com.client as win32 #importou com o  apelido de win32
#importar dados
tabela_vendas =pd.read_excel('vendas.xlsx')

#visualizar os dados
#pd.set_option('display.max_columns', None)
#print(tabela_vendas)

# 2 cochetes e entre parenteses coloca o nome da coluna que ele filtra e aparece so ele


#faturamento por loja

Faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(Faturamento)
print('-' * 50)
#quantidade de produtos vendidos por loja

Qtd_Produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(Qtd_Produtos)
print('-' * 50)
#ticket medio por produto por loja
ticket_medio = (Faturamento['Valor Final']/ Qtd_Produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})#muda o nome da coluna 0 para ticket medio
print(ticket_medio)

#enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'caducarvalho93@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja '
mail.HTMLBody = f'''
<p> Prezados, </p>
<p>Segue o Relatório de Vendas por cada Loja:</p>

<p>Faturamneto:</p>
{Faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade  vendida:</p>
{Qtd_Produtos.to_html()}

<p>Ticket Médio:</p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}

<p>Quaisquer duvida, entre em contato por este email</p> 
<p>Att., Carlos Eduardo</p> 
'''
mail.Send()
print("mandei essa desgraça")

