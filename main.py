import pyautogui
import pyperclip
import time
import pandas
import openpyxl

pyautogui.PAUSE = 1

# dar alt esc do pycharm e entrar no google, entrar no link do google drive

pyautogui.hotkey("alt", "esc")
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing") # para copiar o link
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Navegar no sistema e encontrar a base de vendas
pyautogui.click(x=527, y=358, clicks=2)
time.sleep(2)

# Fazer download da base de vendas

pyautogui.click(x=557, y=357)
pyautogui.click(x=1664, y=236)
time.sleep(2)
pyautogui.click(x=1379, y=780)
time.sleep(5)
pyautogui.click(x=840, y=558)
time.sleep(5)


# Importar para o python a base de vendas e criar as variações
tabela = pandas.read_excel(r"C:\Users\FERNANDO\Downloads\Vendas - Dez.xlsx", engine='openpyxl')
print(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


# Enviar o e-mail

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=282, y=245)
pyautogui.write("febibilemos@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab") # pular para o campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # pular para o campo de corpo do email
texto = f"""Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Fernando Lemos
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")











