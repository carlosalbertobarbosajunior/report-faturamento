#!/usr/bin/env python
# coding: utf-8

# In[1]:


from datetime import date, timedelta
import datetime
from dateutil.relativedelta import relativedelta
import time
import csv
import holidays
import win32com.client
import pandas as pd
import numpy as np
from babel.numbers import format_currency
import sys
import pyodbc
import configparser
import warnings
warnings.filterwarnings("ignore")


# In[2]:


def is_first_business_day():
    '''
    Output:
      today == first_day       - Boolean que identifica se hoje é
                                 o primeiro dia útil do mês
    
    Ação:
      Calcula o primeiro dia útil do mês vigente e o compara com o
      dia de hoje.
    '''
    # lista de feriados do estado do Espírito Santo
    es_holidays = holidays.Brazil(state='ES')
    today = date.today()
    first_day = date(today.year, today.month, 1)
    # se o primeiro dia do mês for um sábado ou domingo, adiciona dias até chegar em um dia útil
    while first_day.weekday() > 4:
        first_day += timedelta(days=1)
    # se o primeiro dia útil for um feriado, adiciona dias até chegar no próximo dia útil
    while first_day in es_holidays:
        first_day += timedelta(days=1)
    return today == first_day

def create_df_from_database(config_file, query):
    '''
    Inputs:
      config_file - Caminho do arquivo .ini (string) que possui as 
                    configurações de conexão com o banco de dados 
                    da empresa.
      query       - Instrução em SQL (string) que será transformada
                    em um dataframe do pandas.
                    
    Outputs:
      df          - Dataframe do pandas construído a partir da instrução
                    SQL da variável query.
    '''
    # executando o configparser e extraindo informações do config.ini
    config = configparser.ConfigParser()
    config.read(config_file)
    server = config['database']['server']
    database = config['database']['database']
    username = config['database']['username']
    password = config['database']['password']
    
    # conectando ao banco de dados e criando o dataframe com a instrução sql
    conn = pyodbc.connect('DRIVER={SQL Server Native Client 10.0};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password)
    df = pd.read_sql(query, conn)
    
    # fechando a conexão para liberá-la
    conn.close()
    
    return df

def format_cell(valor):
    '''
    Input:
      valor          - células do dataframe
      
    Output:
      Múltiplos. Aplica formatação condicional para cada célula
      baseada no valor da mesma.
    '''
    
    if isinstance(valor, str):
        if valor == 'Cancelada':
            return 'background-color: red'
        elif valor.startswith('Devolvida'):
            return 'background-color: yellow'
    return 'border: 1px solid black'


# In[12]:


print('Report de Faturamento')
print('Este script envia dados do faturamento do mês para os usuários selecionados.')
print('Em caso de dúvidas, acesse a documentação em https://github.com/carlosalbertobarbosajunior/report-faturamento')
emails = input('Por favor, digite os e-mails que receberão o report, separados por ponto e vírgula ";"\n')
print('Aguarde enquanto o report é enviado.')

# Caso hoje seja o primeiro dia útil do mês
if is_first_business_day():
    # Considera o início do mês anterior
    first_business_day = date.today() - relativedelta(months=1)
    first_business_day = str(date(first_business_day.year, first_business_day.month, 1))
# Caso contrário
else:
    # Considera o início do mês atual
    first_business_day = date.today()
    first_business_day = str(date(first_business_day.year, first_business_day.month, 1))

# Extraindo a consulta das notas fiscais
faturamento = create_df_from_database(r'\\srvhkm001\GESTAO DE CONTRATOS\Ciência de Dados\Segurança\config.ini', 
                                      'select * from vwNFsProjetos'
                                     )

# Considerando apenas notas fiscais de saída
faturamento = faturamento[faturamento['Saída'] == 'X']

# Reorganizando apenas as colunas necessárias
faturamento = faturamento[['Tipo','NF','Data Emissão','Status','OS','Produto','Cliente','Valor']]

# Convertendo a coluna de data para seu formato correto
faturamento['Data Emissão'] = pd.to_datetime(faturamento['Data Emissão'])

# Reorganizando o dataframe em ordem crescente de data de emissão
faturamento = faturamento.sort_values(by='Data Emissão', ascending=False)

# Filtrando as informações apenas a partir do início do mês
faturamento = faturamento[(faturamento['Data Emissão'] >= first_business_day)]

# Convertendo a data para o formato brasileiro usual
faturamento['Data Emissão'] = faturamento['Data Emissão'].dt.strftime('%d/%m/%Y')

# Extraindo o valor escalar de faturamento, apenas notas autorizadas
vl_faturamento_mensal = faturamento[faturamento['Status']=='Autorizada']['Valor'].sum()

# Convertendo o valor escalar para formato de moeda
vl_faturamento_mensal = 'R${:,.2f}'.format(vl_faturamento_mensal)
# Convertendo as células do dataframe para formato de moeda
faturamento['Valor'] = faturamento['Valor'].apply(lambda x: format_currency(x, currency='BRL', locale='pt_BR'))

# Resetando o índice
faturamento.reset_index(drop=True, inplace=True)

# Aplicar a formatação condicional a cada célula do DataFrame
faturamento = faturamento.style.applymap(format_cell)

# Convertendo para HTML
html = faturamento.render()
# Removendo campos que indicam vazio
html = html.replace('None', '')
html = html.replace('nan', '')

# Enviando e-mail
outlook = win32com.client.Dispatch("Outlook.Application")
Msg = outlook.CreateItem(0)
Msg.To = emails
Msg.Subject = f"Faturamento {date.today().day}-{date.today().month}-{date.today().year}"
Msg.HTMLBody = f'''
 Bom dia,<br><br>
 O faturamento do mês está em <b>{vl_faturamento_mensal}</b>.<br>
 Abaixo, a lista detalhada:<br>
 
 {html}

 <br>OBS: Notas fiscais canceladas ou devolvidas <b>NÃO</b> contabilizam faturamento.<br>
 <br><br>Em caso de dúvidas ou sugestões, favor entrar em contato.<br>
 Este é um e-mail automático, mas sinta-se livre para respondê-lo.
 '''
Msg.Send()

fim = input('Fim do script. Pressione enter para finalizar.')

