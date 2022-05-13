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

# In[64]:


#!pip install pyautogui
#!pip install pyperclip


# In[65]:


# Passo 1 : entrar no sistema da empresaa (no nosso caso vai ser o link do driver)
# Passo 2 : Navegar no sistema e encontrar a base de dados (entrar na pasta exportar )
# Passo 3 : exportar/ fazer donwload da base de dados 
# Passo 4 : importar a base de dados para python
# passo 5 : calcular os indicadores
#passo 6 : enviar um email para diretoria com relatorio


# In[ ]:





# In[66]:


### Vamos agora ler o arquivo baixado para pegar os indicadores


# # import pyautogui
# import pyperclip
# import time

# In[67]:


# Passo 1 : entrar no sistema da empresaa (no nosso caso vai ser o link do driver)
pyautogui.PAUSE = 1
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

#demora alguns segundos 
time.sleep(5)

# Passo 2 : Navegar no sistema e encontrar a base de dados (entrar na pasta exportar )
pyautogui.click(x=428, y=300 ,)
pyautogui.click(clicks=2)

# Passo 3 : exportar/ fazer donwload da base de dados 
pyautogui.click(x=384, y=432  ) #clicar no arquivo
pyautogui.click(x=1722, y=192 ) #clicar nos 3 pontinho 
pyautogui.click(x=1504, y=560 ) #clicar no fazer download
time.sleep(5) #esperar o download se concluir


# In[68]:


# Passo 4 : importar a base de dados para python

import pandas as pd
tabela = pd.read_excel(r"D:\Cursos\#Programação\Aula 1/Vendas - Dez.xlsx")
display(tabela)


# In[69]:


# passo 5 : calcular os indicadores
#faturamento- soma da coluna valor final
faturamento = tabela ["Valor Final"].sum()

#quantidade - soma da venda dos produtos 
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)


# ### Vamos agora enviar um e-mail pelo gmail

# In[70]:


#passo 6 : enviar um email para diretoria com relatorio

import pyautogui
import pyperclip

pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(4)

#clicar no escrever para mandar o relatorio
pyautogui.click(x=74, y=202 )
time.sleep(10)
#escrever o destinatario 
pyautogui.write("xfakeclash@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
#passar para o asunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
#escrever o texto
texto = f"""
    prezados companheiro , bom dia 
    o  faturamento de ontem foi de :R${faturamento:,.2f}
    e a quantidade vendida foi de : {quantidade:,}
    
    abs
    Guilherme"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")

#enviar o email
pyautogui.hotkey("ctrl","enter")


# In[ ]:





# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[71]:


# Passo 2 : Navegar no sistema e encontrar a base de dados (entrar na pasta exportar )pyautogui.position()time.sleep(5)
import time
time.sleep(5)
pyautogui.position()


# In[ ]:





# In[ ]:




