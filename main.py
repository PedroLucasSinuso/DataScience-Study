#Projeto de automação e envio de dados a partir de uma data base aleatória gerada
#Desafios:
# 1. Importar a base de dados
# 2. Calcular o faturamento por loja
# 3. Calcular a quantidade de produtos vendidos por loja
# 4. Calcular o ticket médio por produto em cada loja
# 5. Enviar o relatório por email

import pandas as pd
import win32com.client as win32

#region Importando base de dados

sales_table = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns',None)

#endregion

#region Calculando e exibindo o faturamento por loja

faturamento = sales_table[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#endregion

#region Calculando e exibindo a quantidade de produtos vendidos por loja

qtd_vendidos = sales_table[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(qtd_vendidos)

#endregion

#region Calculando e exibindo o ticket medio por produto em cada loja

ticket_medio = (faturamento['Valor Final']/qtd_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})
print(ticket_medio)

#endregion

#region Enviando o relatório por email


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = '@gmail.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Prezados,</p>
<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:.2f}'.format})}

<p>Quantidade Vendida:</p>
{qtd_vendidos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio':'R${:.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>
<p>Att,</p>
<p>Me</p>
'''
mail.Send()


#endregion