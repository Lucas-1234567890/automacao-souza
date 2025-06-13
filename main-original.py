import warnings
from time import sleep, time
import pandas as pd
import pyautogui
import os
import pyperclip
from tkinter import messagebox, Tk
from openpyxl import load_workbook

# Pop-ups visuais usando tkinter
Tk().withdraw()
messagebox.showinfo("Início da Automação", "A automação foi iniciada com sucesso!")

# Funções auxiliares
def encontrar_imagem(imagem):
    timeout = 20
    inicio = time()
    encontrou = None
    while True:
        try:
            encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.8)
            if encontrou:
                break
        except Exception:
            pass
        if time() - inicio > timeout:
            print(f'Tempo limite atingido. Imagem não encontrada: {imagem}')
            break
        sleep(1)
    return encontrou

def direita(posicoes_imagem):
    return posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3]/2

def esquerda(posicao_imagem, deslocamento=5):
    return posicao_imagem[0] + deslocamento, posicao_imagem[1] + posicao_imagem[3] / 2

def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')

warnings.simplefilter("ignore", UserWarning)

# Leitura da planilha
tabela = pd.read_excel(
    r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\Auto_Elétrica_Souza_Geradores.xlsm",
    sheet_name="Cadastro de materiais",
    skiprows=3,
    usecols="F:L"
)
tabela["Status"] = "Nao"

# Início do processo
pyautogui.FAILSAFE = True
os.startfile(r"C:\Users\Lucas\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\sic.lnk")
sleep(3)
pyautogui.write("123456")
pyautogui.press("tab")
sleep(0.5)
pyautogui.press("enter")
sleep(0.5)
pyautogui.press('enter')
sleep(1)
pyautogui.hotkey("ctrl", "e")
sleep(0.5)
pyautogui.click(pyautogui.center(encontrar_imagem('atualizacao.png')))
sleep(0.5)
pyautogui.click(pyautogui.center(encontrar_imagem('saida.png')))
sleep(0.5)
pyautogui.click(esquerda(encontrar_imagem('outros.png')))
sleep(0.5)
pyautogui.click(pyautogui.center(encontrar_imagem('souza.png')))

# Processar por grupo
grupos = tabela.groupby(["Gerador", "Data"])

for (gerador, data), grupo in grupos:
    try:
        pyautogui.click(direita(encontrar_imagem('gerador.png')))
        pyautogui.write(str(gerador))
        sleep(0.8)
        pyautogui.press('enter')
        sleep(1.5)

        pyautogui.click(pyautogui.center(encontrar_imagem('data.png')))
        escrever_texto(str(data))
        sleep(1.5)

        pyautogui.click(pyautogui.center(encontrar_imagem('quantidade.png')))
        sleep(0.5)

        for idx, linha in grupo.iterrows():
            try:
                escrever_texto(str(linha["Quantidade"]))
                sleep(0.3)
                pyautogui.press('tab')
                sleep(0.3)
                pyautogui.write("00" + str(linha["ID Interno"]))
                sleep(0.3)
                pyautogui.press('tab')
                sleep(0.3)
                pyautogui.press('tab')
                sleep(0.3)
                pyautogui.press('tab')
                sleep(0.3)
                pyautogui.click(1010, 617)  # fecha o pop-up se aparecer
                sleep(0.3)

                tabela.at[idx, "Status"] = "Sim"

            except Exception as e:
                print(f"Erro ao cadastrar linha: {e}")
                tabela.at[idx, "Status"] = "Nao"

        pyautogui.click(pyautogui.center(encontrar_imagem('gravar.png')))
        sleep(1.5)
        pyautogui.click(pyautogui.center(encontrar_imagem('souza.png')))
        sleep(0.8)

    except Exception as e:
        print(f"Erro no grupo ({gerador}, {data}): {e}")

# Exportar resultados
saida_path = r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\log_resultado_automacao.xlsx"

# Adiciona data/hora da execução aos dados novos
tabela["Data Registro"] = pd.Timestamp.now()

# Verifica se o arquivo já existe
if os.path.exists(saida_path):
    # Lê o conteúdo antigo
    tabela_antiga = pd.read_excel(saida_path)
    # Junta com os dados novos
    tabela_final = pd.concat([tabela_antiga, tabela], ignore_index=True)
else:
    tabela_final = tabela

# Salva a nova versão completa (sem sobrescrever dados anteriores)
tabela_final.to_excel(saida_path, index=False)

# Pop-up final
messagebox.showinfo("Automação Finalizada", f"A automação foi concluída!\n\nLog salvo em:\n{saida_path}")
