import pyautogui
import pyperclip
import time
import pandas as pd

#pyautogui.click > clicar
#pyautogui.write > escrever
#pyautogui.press > presssionar tecla
#pyautogui.hotkey > atalhos (alt + tab e etc)
#Documentação > https://pyautogui.readthedocs.io/en/latest/


pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

# Passo 1: Entrar no sistema da empresa (lnk do drive)
pyautogui.press("win")
pyautogui.write("opera")
pyautogui.press("enter")
time.sleep(2)
pyautogui.hotkey('ctrl'+'l')
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")

# Passo 2: Navegar até local do relatório
time.sleep(2)
pyautogui.click(x=423, y=294, clicks=2)

#Passo 3: Exportar o relatório
time.sleep(2)
pyautogui.click(x=423, y=294)
pyautogui.click(x=1711, y=184)
pyautogui.click(x=1544, y=593)
time.sleep(5)
pyautogui.press('enter')

#Passo 4: Calcular os indicadores (faturamento e quantidade de produtos)
tabela = pd.read_excel(r'C:\Users\Gustavo\Documents\Vendas - Dez.xlsx')
faturamento = tabela["Valor Final"].sum()
quantidade = tabela['Quantidade'].sum()

faturamento = f'{faturamento:_.2f}'
faturamento = faturamento.replace('.',',')
faturamento = faturamento.replace('_','.')
quantidade = f'{quantidade:_}'
quantidade = quantidade.replace('_','.')
print(faturamento)
print(quantidade)

#Passo 5: Enviar análise via e-mail
time.sleep(2)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")

# botão escrever
time.sleep(3)
pyautogui.click(x=132, y=213)
time.sleep(1)
pyautogui.write(r'gustavo.borba70@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
#pyautogui.click(x=1287, y=525)
pyperclip.copy('Relátorio de vendas')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')

texto = f"""Prezados,
 
Segue faturamento mensal e dados em anexo.
O faturamento foi de: R${faturamento}.
A quantidade de itens vendidos foi de: {quantidade} unidades.


Atenciosamente, 
Gustavo Barbosa
"""

pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')

#botão anexar
pyautogui.click(x=1426, y=1007)

#selecionar arquivo
pyautogui.click(x=1498, y=821)
time.sleep(3)

#enviar email
pyautogui.hotkey('ctrl','enter')

#botao enviar (x=1301, y=1002)
'''time.sleep(5)
print(pyautogui.position())'''







'''
# Mecanismo para identificar a posição desejada do "mouse"
time.sleep(5)
print(pyautogui.position())'''




