import pandas as pd
import urllib
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from datetime import datetime


# CONFIGURAÇÕES
CAMINHO_ARQUIVO = r"C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\Auto_Elétrica_Souza_Geradores.xlsm"
ABA = "Preventiva"
TELEFONES = ["71999365938"]  

def gerar_mensagem_whatsapp(df):
    linhas = []

    # Saudações inteligentes
    hora = datetime.now().hour
    if hora < 12:
        saudacao = "Bom dia"
    elif hora < 18:
        saudacao = "Boa tarde"
    else:
        saudacao = "Boa noite"
    
    linhas.append(f"{saudacao}, Rodrigo.")
    linhas.append("📋 *Relatório de Manutenções Preventivas*")
    linhas.append("Legenda: ✅ Em dia | 🔵 Amanhã | 🟡 Hoje | 🔴 Vencido\n")

    for _, row in df.iterrows():
        situacao = row['Situação'].strip().lower()
        if "vencido" in situacao:
            emoji = "🔴"
        elif "hoje" in situacao:
            emoji = "🟡"
        elif "amanhã" in situacao or "amanha" in situacao:
            emoji = "🔵"
        else:
            emoji = "✅"

        linha = f"{emoji} *{row['Geradores']}* — {row['Localização']} — {row['data do prazo'].strftime('%d/%m/%Y')} — _{row['Situação']}_"
        linhas.append(linha)

    linhas.append("\n⏰ Verifique o status das manutenções e programe-se.")
    linhas.append("Abraços, Equipe de Manutenção 👷‍♂️")
    return "\n".join(linhas)


# LÊ O EXCEL E PREPARA O DATAFRAME
df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=ABA, usecols="G:S", skiprows=5)
df = df[["Geradores", "Localização", "data do prazo", "Situação"]]
df["data do prazo"] = pd.to_datetime(df["data do prazo"], errors='coerce')
df = df[df["data do prazo"].notna()]

# GERA A MENSAGEM
mensagem = gerar_mensagem_whatsapp(df)

# ABRE O WHATSAPP WEB
servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument(r"user-data-dir=C:\Users\Lucas\AppData\Local\Temp\Profile Selenium")  # Mantém login
navegador = webdriver.Chrome(service=servico, options=options)

navegador.get("https://web.whatsapp.com")
print("🔓 Faça login no WhatsApp Web...")

while len(navegador.find_elements(By.ID, 'side')) < 1:
    sleep(1)
sleep(2)

# ENVIA A MENSAGEM PARA CADA TELEFONE
for telefone in TELEFONES:
    texto = urllib.parse.quote(mensagem)
    link = f"https://web.whatsapp.com/send?phone=55{telefone}&text={texto}"
    navegador.get(link)

    try:
        wait = WebDriverWait(navegador, 15)
        wait.until(EC.presence_of_element_located((By.ID, 'side')))
        botao_enviar = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[4]/button')
        ))
        botao_enviar.click()
        print(f"✅ Mensagem enviada para {telefone}")
        sleep(5)
    except Exception as e:
        print(f"❌ Falha ao enviar para {telefone}: {e}")
        sleep(5)

navegador.quit()
print("🚀 Fim do envio!")
