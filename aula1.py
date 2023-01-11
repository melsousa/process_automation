import pyautogui
import pyperclip
import time
import pandas as pd


pyautogui.PAUSE = 3


# passo 1: Entrar no sistema da empresa

pyautogui.press("win")
pyautogui.write("firefox")
pyautogui.press("enter")
pyautogui.write(
    "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.press("enter")

time.sleep(5)

# passo 2: Navegar até a local do relatório

pyautogui.click(x=441, y=291, clicks=2)
time.sleep(4)

# passo 3: Exportar o relatório

pyautogui.click(x=377, y=398)
pyautogui.click(x=1154, y=168)
pyautogui.click(x=969, y=572)
time.sleep(5)

print(pyautogui.position())

# passo 4: Calcular os indicadores

tabela = pd.read_excel(r"D:\Users\Acer\Downloads\Vendas - Dez.xlsx")


faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


# passo 5: Enviar um email para a diretoria

# abrir aba e entrar no email
pyautogui.press("win")
pyautogui.write("firefox")
pyautogui.press("enter")
pyautogui.write(
    "https://mail.google.com/mail/u/2/#inbox")
pyautogui.press("enter")

time.sleep(10)

# clicar no botão escrever
pyautogui.click(x=110, y=176)
time.sleep(2)

# preencher as informações do email
pyautogui.click(x=985, y=300)
pyautogui.write("melissasousa3ms@gmail.com")
pyautogui.press("tab")  # seleciona o email
pyautogui.press("tab")  # pula pro campo assunto

# assunto
time.sleep(1)
pyautogui.write("Relatório de Vendas")
pyautogui.press("tab")

# corpo
texto = f""" 
    Prezados, bom dia
    O faturamento de ontem foi de: {faturamento:,.2f}
    A quantidade de produtos foi de: {quantidade:,}
    
    Abs Melissa
    """
pyautogui.write(texto)

# enviar

pyautogui.hotkey("ctrl", "enter")

# print(pyautogui.position())






# passo 5: Enviar um email para a diretoria
