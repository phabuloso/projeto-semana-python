import pyautogui
import time
import pyperclip

pyautogui.PAUSE = 1
pyautogui.alert("Vai começar, aperte OK e não mexa em nada")

# opção 1 - abrir navegador novo e entrar no chrome
# pyautogui.press("winleft")
# pyautogui.write("chrome")
# pyautogui.press("enter")

# opção 2 - abrir uma nova aba
pyautogui.hotkey('ctrl', 't')

# abrir drive
# ensinar aqui o write
link = "https://docs.google.com/spreadsheets/d/1IzfCMOe37HmsMEbwJOXyttjiVsfkWZfQV46K0peipq0/edit?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# baixar base de dados atualizada
pyautogui.click(935, 694, clicks=2)
pyautogui.click(2028, 895)
pyautogui.click(3306, 406)
pyautogui.click(2880, 1489)
time.sleep(10)

#Vamos agora ler o arquivo baixado para pegar os indicadores
#Faturamento
#Quantidade de Produtos

import pandas as pd

df = pd.read_excel("C://Users/joaop/Downloads/Vendas - Dez.xlsx")
display(df)
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()

# Vamos agora enviar um e-mail pelo gmail

import pyperclip

# abrir aba gmail
pyautogui.hotkey('ctrl', 't')
pyautogui.write("mail.google.com")
pyautogui.press('enter')
time.sleep(5)

# clicar em escrever email
pyautogui.click(307, 506)

# preencher informações do e-mail
pyautogui.write('phablopatriciof@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = "Relatório de Vendas de Ontem"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", 'v')
pyautogui.press("tab")
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,}

Abs
LiraPython"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", 'v')

# enviar e-mail
pyautogui.hotkey('ctrl', 'enter')

# avisar que acabou
pyautogui.alert("Fim da Automação. Seu computador já voltou a ser seu")

#Use esse código para descobrir qual a posição de um item que queira clicar
#Lembre-se: a posição na sua tela é diferente da posição na minha tela

import pyautogui
import time
time.sleep(4)
print(pyautogui.position())
pyautogui.alert("Posição Registrada")

#Caso queira pegar por uma imagem

time.sleep(5)
x, y = pyautogui.locateCenterOnScreen("txt.png")
pyautogui.click(x, y)