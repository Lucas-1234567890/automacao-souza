import warnings
from time import sleep, time
import pandas as pd
import pyautogui
import os
import pyperclip
from tkinter import Tk, Label, Button, filedialog, messagebox
from openpyxl import load_workbook
import threading
from datetime import datetime

# ---------- Funções auxiliares ----------
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
    return posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3] / 2

def esquerda(posicao_imagem, deslocamento=5):
    return posicao_imagem[0] + deslocamento, posicao_imagem[1] + posicao_imagem[3] / 2

def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')

# ---------- Função Principal da Automação ----------
def iniciar_automacao(arquivo_excel):
    try:
        tabela = pd.read_excel(arquivo_excel, sheet_name="Cadastro de materiais", skiprows=3, usecols="F:L")
        tabela["Data"] = pd.to_datetime(tabela["Data"]).dt.strftime("%d/%m/%Y")
        tabela["Status"] = "Nao"

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

        pyautogui.click(pyautogui.center(encontrar_imagem(os.path.join('imagens', 'atualizacao.png'))))
        sleep(0.5)
        pyautogui.click(pyautogui.center(encontrar_imagem(os.path.join('imagens', 'saida.png'))))
        sleep(0.5)
        pyautogui.click(esquerda(encontrar_imagem(os.path.join('imagens', 'outros.png'))))
        sleep(0.5)
        pyautogui.click(pyautogui.center(encontrar_imagem(os.path.join('imagens', 'max.png'))))
        sleep(1)

        grupos = tabela.groupby(["Gerador", "Data"])

        for (gerador, data), grupo in grupos:
            try:
                pyautogui.click(pyautogui.center(encontrar_imagem(os.path.join('imagens', 'souza.png'))))
                sleep(0.8)
                pyautogui.click(direita(encontrar_imagem(os.path.join('imagens', 'gerador.png'))))
                pyautogui.write(str(gerador))
                sleep(0.8)
                pyautogui.press('enter')
                sleep(1.5)
                pyautogui.doubleClick(79, 60)
                sleep(1)
                escrever_texto(str(data))
                sleep(1.5)

                pyautogui.click(pyautogui.center(encontrar_imagem(os.path.join('imagens', 'quantidade.png'))))
                sleep(0.5)

                for idx, linha in grupo.iterrows():
                    try:
                        escrever_texto(str(linha["Quantidade"]))
                        sleep(0.3)
                        pyautogui.press('tab')
                        sleep(0.3)
                        pyautogui.write(str(linha["ID Interno"]).zfill(6))
                        sleep(0.3)
                        pyautogui.press('tab')
                        sleep(0.3)
                        pyautogui.press('tab')
                        sleep(0.3)
                        pyautogui.press('tab')
                        sleep(1)
                        imagem = pyautogui.locateCenterOnScreen(os.path.join('imagens', 'sim.png'), confidence=0.9)
                        if imagem:
                            pyautogui.click(imagem.x, imagem.y)
                        sleep(0.3)
                        tabela.at[idx, "Status"] = "Sim"

                    except Exception as e:
                        print(f"Erro ao cadastrar linha: {e}")
                        tabela.at[idx, "Status"] = "Nao"

                sleep(1)
                pyautogui.click(pyautogui.center(encontrar_imagem(os.path.join('imagens', 'gravar.png'))))
                sleep(1.5)
                imagem_2 = pyautogui.locateOnScreen(os.path.join('imagens', 'sim_nao.png'), confidence=0.9)
                sleep(0.3)
                if imagem_2:
                    sleep(1)
                    pyautogui.click(esquerda(imagem_2))
                sleep(1)
        
            except Exception as e:
                print(f"Erro no grupo ({gerador}, {data}): {e}")

        pasta_logs = r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\logs_automacao"
        os.makedirs(pasta_logs, exist_ok=True)

        data_hora_execucao = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        saida_path = os.path.join(pasta_logs, f"log_{data_hora_execucao}.xlsx")

        tabela["Data Registro"] = pd.Timestamp.now()
        tabela.to_excel(saida_path, index=False)

        messagebox.showinfo("Concluído", f"A automação foi finalizada!\n\nLog salvo em:\n{saida_path}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

# ---------- Interface Tkinter ----------
def selecionar_arquivo():
    caminho = filedialog.askopenfilename(title="Selecione a planilha Excel", filetypes=[("Planilhas Excel", "*.xls*")])
    if caminho:
        btn_iniciar["state"] = "normal"
        lbl_caminho.config(text=f"Arquivo selecionado:\n{caminho}")
        btn_iniciar.config(command=lambda: threading.Thread(target=iniciar_automacao, args=(caminho,)).start())

root = Tk()
root.title("Automação de Cadastro - Auto Elétrica Souza")
root.geometry("500x250")

Label(root, text="Automação de Cadastro no SIC", font=("Arial", 16, "bold")).pack(pady=10)
Button(root, text="Selecionar Planilha", command=selecionar_arquivo).pack(pady=5)

lbl_caminho = Label(root, text="Nenhum arquivo selecionado", fg="gray")
lbl_caminho.pack(pady=5)

btn_iniciar = Button(root, text="Iniciar Automação", state="disabled")
btn_iniciar.pack(pady=10)

Button(root, text="Sair", command=root.quit).pack(pady=10)

root.mainloop()