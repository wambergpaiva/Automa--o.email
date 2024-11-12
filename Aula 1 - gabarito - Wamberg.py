#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[3]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1


 
    
# passo 1: Entrar no sistema (no nosso caso entrar no link)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(5)
# passo 2: Navegar até local do relatório (entrar na pasta exportar)

pyautogui.click(x=391, y=339, clicks=2 )
time.sleep(2)

# passo 3: Fazer o dowload do relatório

pyautogui.click(x=362, y=439)
pyautogui.click(x=606, y=206)
pyautogui.click(x=678, y=282)

time.sleep(5)   



# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[15]:


# passo 4: Calcular os indicadores

import pandas as pd

tabela = pd.read_excel(r"C:\Users\wambe\Downloads\Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()




# ### Vamos agora enviar um e-mail pelo gmail

# In[5]:


# passo 5: Entrar no email

pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(5)

# passo 6: Enviar por e-mail os resultados
pyautogui.click(x=60, y=144)

pyautogui.write('tricolorflu34+diretoria@gmail.com')
pyautogui.press('tab') # seleciona o email
# Para mandar mais emails
# escreve outro email
# tab
# escreve outro email
# tab
pyautogui.press('tab') # pula pro campo do assunto
pyperclip.copy('Relatório de vendas') # escrever o assunto - Sempre que tiver caracter especial, tem que copiar e colar.
pyautogui.hotkey('ctrl', 'v') # escrever o assunto
pyautogui.press('tab') # pular pro corpo do emal

texto = f"""

Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}
    
Abs
Wamberg Paiva - Cientista de Dados """

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

# Clicar no botão enviar
# Apertar ctrl + enter

pyautogui.hotkey('ctrl', 'enter')


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[6]:


pyautogui.sleep(5)
pyautogui.position() 


# In[7]:


# Como instalar

get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# In[ ]:





# In[ ]:





# In[ ]:




