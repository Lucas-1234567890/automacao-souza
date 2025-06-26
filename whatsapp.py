import pandas as pd
import urllib
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep

# CONFIG
CAMINHO_ARQUIVO = r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\Auto_Elétrica_Souza_Geradores.xlsm"
ABA = "Cadastro de materiais"
CAMINHO_LOG = r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\log_envios.csv"

# CARREGAR PLANILHA
tabela = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=ABA, skiprows=3, usecols="F:N")
tabela["Data"] = pd.to_datetime(tabela["Data"]).dt.strftime("%d/%m/%Y")

# CARREGAR LOG DE ENVIOS (ou criar vazio se não existir)
if os.path.exists(CAMINHO_LOG):
    log_envios = pd.read_csv(CAMINHO_LOG, dtype=str)
else:
    log_envios = pd.DataFrame(columns=["Técnico", "Telefone", "Gerador", "Data"])

mensagens = []

# AGRUPAR POR Técnico, Telefone, Gerador e Data
grupos = tabela.groupby(["Técnico", "Telefone", "Gerador", "Data"])

for (tecnico, telefone, gerador, data), grupo in grupos:
    if pd.isna(telefone):
        continue

    # Verificar se já foi enviado
    filtro = (
        (log_envios["Técnico"] == tecnico) &
        (log_envios["Telefone"] == str(int(telefone))) &
        (log_envios["Gerador"] == gerador) &
        (log_envios["Data"] == data)
    )

    if filtro.any():
        continue  # já foi enviado

    # Montar mensagem
    msg = f"👷 *{tecnico}*, tudo certo?\n\n"
    msg += f"📅 _Relatório de uso de materiais - {data}_\n"
    msg += f"⚙️ *Gerador:* _{gerador}_\n\n"
    msg += f"*Materiais utilizados:*\n"

    for _, linha in grupo.iterrows():
        nome = linha.get("Materiais") or "Sem nome"
        qtde = linha.get("Quantidade") or 0
        msg += f"• *{nome}* — _{qtde} unidade(s)_\n"

    msg += "\n✅ Por favor, confirme o uso ou reporte qualquer divergência.\n"
    msg += "Obrigado! 😊"

    mensagens.append((str(int(telefone)), msg, tecnico, gerador, data))

# ABRIR CHROME COM PERFIL
servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument(r"user-data-dir=C:\Users\Lucas\AppData\Local\Temp\Profile Selenium")
navegador = webdriver.Chrome(service=servico, options=options)

navegador.get("https://web.whatsapp.com")
print("🔓 Faça login no WhatsApp Web...")
while len(navegador.find_elements(By.ID, 'side')) < 1:
    sleep(1)
sleep(2)

# ENVIAR AS MENSAGENS E ATUALIZAR LOG
for telefone, mensagem, tecnico, gerador, data in mensagens:
    texto = urllib.parse.quote(mensagem)
    link = f"https://web.whatsapp.com/send?phone=55{telefone}&text={texto}"

    navegador.get(link)
    while len(navegador.find_elements(By.ID, 'side')) < 1:
        sleep(1)
    sleep(3)

    try:
        botao_enviar = navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[4]/button')
        botao_enviar.click()
        print(f"✅ Mensagem enviada para {telefone}")

        # Salvar no log
        novo_log = pd.DataFrame([{
            "Técnico": tecnico,
            "Telefone": telefone,
            "Gerador": gerador,
            "Data": data
        }])
        log_envios = pd.concat([log_envios, novo_log], ignore_index=True)
        log_envios.to_csv(CAMINHO_LOG, index=False)

        sleep(5)
    except Exception as e:
        print(f"❌ Falha ao enviar para {telefone}:", e)

navegador.quit()
