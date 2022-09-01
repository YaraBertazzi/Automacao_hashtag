import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

#abrir o navegador
pyautogui.press('winleft')
pyautogui.write ('chrome')
pyautogui.press('enter')
pyautogui.alert('Vai começar, aperte ok e não mexa em nada')
#pyautogui.click(x=824, y=409, clicks=2)  # clicla na pasta exportar
time.sleep(2)

pyautogui.hotkey('ctrl', 't')


# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(15)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=421, y=261, clicks=2) #clicla na pasta exportar
time.sleep(5)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=315, y=379) #clicar no https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing

pyautogui.click(x=553, y=171) #clicar no 3 pontinhos
pyautogui.click(x=363, y=600) #clicar em download
#time.sleep(5)

# Passo 4: Calcular os indicadores

import pandas as pd  #pd é um apaledido para pandas

tabela = pd.read_excel(r'C:\Users\HD\Downloads/Vendas - Dez.xlsx') #endereço da planilha do excel para importar
print(tabela)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

# Passo 5: Entrar no email
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(10)

#https://mail.google.com/mail/u/0/#inbox

# '''# Passo 6: Enviar por e-mail o resultado

pyautogui.click(x=91, y=183) #arrumar
time.sleep(7)

pyautogui.write('yarabertazzi@gmail.com')
pyautogui.press('tab') # seleciona o email
# escreve outro email
# tab
# escreve outro email
# tab
pyautogui.press('tab') # pula pro campo de assunto
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl', 'v') # escrever o assunto
pyautogui.press('tab') #pular pro corpo do email

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
LiraPython"""

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

#clicar no botão enviar

#apertar Ctrl Enter
pyautogui.hotkey('ctrl', 'enter')




