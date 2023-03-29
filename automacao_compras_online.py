#!/usr/bin/env python
# coding: utf-8

# # Projeto Automação Web - Busca de Preços
# 
# ### Objetivo: 
# Criar automações web com Selenium para buscar as informações na internet
# 
# 
# ### Desafio:
# 
# - Imagina que você trabalha na área de compras de uma empresa e precisa fazer uma comparação de fornecedores para os seus insumos/produtos.
# 
# - Nessa hora, você vai constantemente buscar nos sites desses fornecedores os produtos disponíveis e o preço, afinal, cada um deles pode fazer promoção em momentos diferentes e com valores diferentes.
# 
# 
# ### Base de dados:
# 
# - Planilha de Produtos, com os nomes dos produtos, o preço máximo, o preço mínimo (para evitar produtos "errados" ou "baratos de mais para ser verdade") e os termos banidos que vamos querer evitar nas nossas buscas.
# 
# ### O que devemos fazer:
# 
# - Procurar cada produto no Google Shopping e pegar todos os resultados corretos e que tenham preço dentro da faixa;
# - Enviar um e-mail para o seu e-mail (no caso da empresa seria para a área de compras por exemplo) com a notificação e a tabela com os itens e preços encontrados, junto com o link de compra.

# In[7]:


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
servico = Service(ChromeDriverManager().install())

import pandas as pd
import win32com.client as win32
from pprint import pprint
from datetime import datetime
import time
import os


# In[2]:


# importar a base de dados
df = pd.read_excel('buscas.xlsx')
display(df)
df.info()


# In[3]:


# acessar o navegador
web = webdriver.Chrome(service=servico)
web.maximize_window()

for produto in df['Nome']:

    # definir os termos
    lista_termos_produto = produto.split(' ')
    termos_banidos = df.loc[df['Nome']==produto, 'Termos banidos']
    try:
        lista_termos_banidos = termos_banidos.split(',')
    except:
        lista_termos_banidos = termos_banidos
    preco_min = df.loc[df['Nome']==produto, 'Preço mínimo']
    if float(preco_min) != '':
        preco_min = float(preco_min)
    preco_max = df.loc[df['Nome']==produto, 'Preço máximo']
    if float(preco_max) != '':
        preco_max = float(preco_max)

    # acessar o google
    web.get('https://www.google.com.br/')
    time.sleep(2)

    # digitar o termo procurado
    web.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto, Keys.ENTER)
    time.sleep(1)
    
    # clicar em shopping
    web.find_element(By.XPATH, '//*[@id="hdtb-msb"]/div[1]/div/div[2]/a').click()

    # percorrer todas os resultados (buscar uma classe/id em comum a todas os resultados)
    lista_resultados = web.find_elements(By.CLASS_NAME, 'sh-dgr__gr-auto')

    # armazenar as informações de cada resultado (nome do produto, valor, link da loja)
    lista = []

    for resultado in lista_resultados:

        #nome
        nome_produto = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        nome_produto = nome_produto.lower().replace('(', '').replace(')', '').replace('|', '').replace(',', '')
        lista_termos_resultado = nome_produto.split(' ')

        for termo in lista_termos_produto:
            if termo in lista_termos_resultado and termo not in lista_termos_banidos:

                #preco
                preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                preco = preco.replace('R$', '').replace(' ','').replace('.','').replace(',', '.').replace('+impostos', '')
                if preco != '':
                    preco = float(preco)

                if float(preco_min) <= preco <= float(preco_max):

                    #link 
                    elemento_filho = resultado.find_element(By.CLASS_NAME, 'aULzUe') #como não foi possível acessar o link da loja diretamente pelo elemento com a url, busca-se um elemento dentro deste (elemento filho) e, a partir dele, procura-se o elemento acima (elemento pai); '..' significa o elemento que está diretamente acima (elemento pai)
                    elemento_pai = elemento_filho.find_element(By.XPATH, '..') 
                    link = elemento_pai.get_attribute('href')

                    #adicionar à lista
                    lista.append((nome_produto, preco, link))    
        
#fechar o navegador
web.quit()

for tupla in lista:
    pprint(tupla)

df.info()


# In[4]:


# adicionar as ofertas encontradas em uma tabela unificada
df_ofertas = pd.DataFrame(columns=['Nome', 'Preço', 'Link'])

n_linha = 0
for nome, preco, link in lista:
    df_ofertas.loc[n_linha, 'Nome'] = nome
    df_ofertas.loc[n_linha, 'Preço'] = preco
    df_ofertas.loc[n_linha, 'Link'] = link
    n_linha += 1

df_ofertas = df_ofertas.drop_duplicates() # remover linhas duplicadas

display(df_ofertas)
df_ofertas.info()


# In[5]:


# exportar o df final
df_ofertas.to_excel('Ofertas Google Shopping (revisão).xlsx')


# In[6]:


# enviar email com o df das ofertas
hoje = datetime.now().strftime('%d/%m/%Y')

if len(df_ofertas.index) > 0: # verificar se há ofertas no dia
    
    outlook = win32.Dispatch('outlook.application')
    e = outlook.CreateItem(0)
    e.To = 'bep_rafael@hotmail.com'
    e.Subject = f'Ofertas do Google Shopping - {hoje}'
    e.HTMLBody = f'''<p>Bom dia, diretoria</p>
<p>Segue tabela com as ofertas encontradas:</p>
<p>{df_ofertas.to_html(index=False)}</p>
<p>Att,</p>
<p>Rafael Muller.</p>
'''
    local = os.getcwd()
    arquivo = fr'{local}\Ofertas Google Shopping (revisão).xlsx')
    e.Attachments.Add(str(arquivo))
    e.Send()

else:
    pass


# In[ ]:




