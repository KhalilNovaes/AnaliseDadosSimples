import pandas as pd;
import win32com.client as win32

emailDestiny = 'fieldlgtestk2020@gmail.com';

tabela_vendas = pd.read_excel('Base de Dados/Vendas.xlsx');

pd.set_option('display.max_columns', None);

print(tabela_vendas);

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum();

print(faturamento);

qtd = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum();

ticket_medio = (faturamento['Valor Final'] / qtd['Quantidade']).to_frame(); 
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})


outlook = win32.Dispatch('outlook.application');
email = outlook.CreateItem(0);
email.To = emailDestiny;
email.Subject = 'Relatório de Vendas';
email.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue abaixo o Relatório de Vendas.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{qtd.to_html()}

<p>Ticket Médio dos Produtos por Loja</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida me encontro á disposição.</p>

<p>Att,</p>
'''

email.Send();