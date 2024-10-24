# HTML-PYTHON
# ANALISE GRAFICA USANDO MACHINE LEARNING E ENVIO DE EMAIL AUTOMATICO
import pandas as pd
import win32com.client as win32

# Importar base de dados
df_vendas = pd.read_excel('C:/Users/suporte/Downloads/Vendas.xlsx')

# Visualizar base de dados
pd.set_option('display.max_columns', None)
print(df_vendas)

# Faturamento por loja
faturamento_por_loja = df_vendas.groupby('ID Loja')['Valor Final'].sum()
print("\nFaturamento por Loja:")
# Ordenado por faturamento decrescente
print(faturamento_por_loja.sort_values(ascending=False))

# Quantidade de produtos vendidos por loja
quantidade_vendida = df_vendas.groupby('ID Loja')['Quantidade'].sum()
print("\nQuantidade de Produtos Vendidos por Loja:")
# Ordenado por quantidade decrescente
print(quantidade_vendida.sort_values(ascending=False))

# Ticket médio por loja
ticket_medio = faturamento_por_loja / quantidade_vendida
print("\nTicket Médio por Loja:")
# Ordenado por ticket médio decrescente
print(ticket_medio.sort_values(ascending=False))

# Enviar email com relatório
try:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'vinicius.x.pirata@gmail.com'  # Substitua pelo endereço de destino
    mail.Subject = 'Relatório de Vendas por Loja'
    mail.HTMLBody = f'''
    <p>Vinizeira,</p>
    <p>Relatório de vendas por cada loja, meninão!</p>

    <p><strong>Faturamento:</strong></p>
    {faturamento_por_loja.to_frame().to_html()}

    <p><strong>Quantidade Vendida:</strong></p>
    {quantidade_vendida.to_frame().to_html()}

    <p><strong>Ticket Médio dos Produtos em Cada Loja:</strong></p>
    {ticket_medio.to_frame().to_html()}

    <p>Qualquer erro no código faz parte.....</p>
    '''

    # Para exibir o e-mail antes de enviar
    mail.Display()

    # Para enviar o e-mail diretamente
    # mail.Send()

    print('E-mail enviado com sucesso!')

except Exception as e:
    print(f'Ocorreu um erro ao enviar o e-mail: {e}')
