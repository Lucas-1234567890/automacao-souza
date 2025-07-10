import pyautogui
import time
import pandas as pd

def esquerda(posicao_imagem, deslocamento=5):
    return posicao_imagem[0] + deslocamento, posicao_imagem[1] + posicao_imagem[3] / 2

tabela = pd.read_excel(
    r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\Auto_Elétrica_Souza_Geradores.xlsm",
    sheet_name="Cadastro de materiais",
    skiprows=3,
    usecols="F:L"
)

tabela["Data"] = pd.to_datetime(tabela["Data"]).dt.strftime("%d/%m/%Y")

print(tabela)
# Agrupamento com contagem


#time.sleep(3)
#posicao = pyautogui.position()
#print(f"Posição do mouse: X={posicao[0]}, Y={posicao[1]}")
#imagem = pyautogui.locateOnScreen('sim_nao.png',confidence=0.9)
#pyautogui.click(esquerda(imagem))