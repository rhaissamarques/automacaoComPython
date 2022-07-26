import pyautogui
import pyperclip
import time
import pandas as pd

#abrir uma nova aba
#copiar o link
#colar o link
#apertar enter
pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(7)

#clicar com o mouse no botão exportar - duplo clique
#clicar no arquivo
#clicar nos 3 pontinhos
#clicar em download
#mude os x e y para o seu monitor, esses funcionam no meu
pyautogui.click(x=424, y=310, clicks=2)
time.sleep(7)
pyautogui.click(x=417, y=374)
time.sleep(1)
pyautogui.click(x=1164, y=188)
time.sleep(1)
pyautogui.click(x=983, y=595)
time.sleep(5)

#importar o arquivo
#tirar as informações que precisamos
#criar um relatório
#coloca o r antes do caminho para interpretar as \ como parte do caminho para o arquivo
#o display só funciona no Jupyter, é a mesma coisa do print, mas a tabela aparece mais bonita

tabela = pd.read_excel(r"C:\Users\admin\Downloads\Vendas - Dez.xlsx")
display(tabela)

#fazendo as contas dos indicadores
quantidade = tabela["Quantidade"].sum()
valorFinal = tabela["Valor Final"].sum()
print(quantidade)

#abrir uma nova guia
#copiar o link da caixa de entrada
#enter
#entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/?ogbl#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)
#clicar no mais para escrever um novo email
#escrever o destinatário
#escrever o assunto do email
#escrever o conteúdo do email
pyautogui.click(x=99, y=199)
time.sleep(5)
pyautogui.write("seuemail@gmail.com")
pyautogui.press("enter")
pyautogui.press("tab")
pyperclip.copy("Análise de indicadores")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
#o f no início é para que os {} funcionem
texto = f"""
Prezados, bom dia.
O faturamento de ontem foi de: R${valorFinal:,.2f}
A quantidade de produtos vendida foi de: {quantidade:,}
Qualquer dúvida estou à disposição.
Att.:
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
print(valorFinal)
#fim, o email chegou na conta que vc escolheu...
#adapte para o que precisar.
#XOXO
