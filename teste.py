import pyautogui
import time
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\Auto_Elétrica_Souza_Geradores.xlsm", sheet_name="Cadastro de materiais", skiprows=3, usecols="F:L")
tabela["Data"] = pd.to_datetime(tabela["Data"]).dt.strftime("%d/%m/%Y")
print(tabela)


# Aguarda 3 segundos para permitir que você posicione o mouse
time.sleep(5)

# Obtém a posição atual do mouse
posicao = pyautogui.position()

# Imprime a posição e encerra o programa
print(f"Posição do mouse: X={posicao[0]}, Y={posicao[1]}")
