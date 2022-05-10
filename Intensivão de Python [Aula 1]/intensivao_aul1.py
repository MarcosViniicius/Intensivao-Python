import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

pyautogui.click(x=375, y=270, clicks=2)
time.sleep(3)
pyautogui.click(x=379, y=419)
time.sleep(3)
pyautogui.click(x=1169, y=171)
time.sleep(3)
pyautogui.click(x=960, y=559)

time.sleep(5) # Espera de download

tabela = pd.read_excel(r"C:\Users\marco\Downloads\Vendas - Dez.xlsx")

faturamento = tabela["Valor Final"].sum()

quantidade = tabela["Quantidade"].sum()

print(quantidade)
print(faturamento)

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(4)

pyautogui.click(x=90, y=167)

time.sleep(2)

pyautogui.write("marcosvinicius.lucena28@gmail.com")
pyautogui.press("tab") # selecionando o destinatário
pyautogui.press("tab") # passando para o campo de assunto
pyperclip.copy("Relatório de vendas.")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab")

texto = f"""
Prezados, bom dia!

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: R${quantidade:,}

Abs
Marcos Vinicius"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

time.sleep(3)

pyautogui.hotkey("ctrl", "enter")