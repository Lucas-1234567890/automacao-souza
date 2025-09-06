import keyboard
import os

# Caminhos dos scripts
SCRIPT_WHATSAPP = r"C:\Desktop\automacao-souza-main\whatsapp.py"
SCRIPT_PREVENTIVA = r"C:\Desktop\automacao-souza-main\preventiva.py"
SCRIPT_AUTOMACAO = r"C:\Desktop\automacao-souza-main\main.py"
SCRIPT_ENTRADA_SAIDA = r'C:\Desktop\automacao-souza-main\entrada_saida.py'

def rodar_whatsapp():
    print("🚀 Executando script de envio de WhatsApp...")
    os.system(f'python "{SCRIPT_WHATSAPP}"')

def rodar_preventiva():
    print("🚀 Executando script de manutenção preventiva...")
    os.system(f'python "{SCRIPT_PREVENTIVA}"')

def rodar_automacao():
    print("🚀 Executando script de automação...")
    os.system(f'python "{SCRIPT_AUTOMACAO}"')

def rodar_entrada_saida():
    print("🚀 Executando script de entrada e saida...")
    os.system(f'python "{SCRIPT_ENTRADA_SAIDA}"')

# Atalhos
keyboard.add_hotkey("ctrl+alt+w", rodar_whatsapp)
keyboard.add_hotkey("ctrl+alt+p", rodar_preventiva)
keyboard.add_hotkey("ctrl+alt+a", rodar_automacao)
keyboard.add_hotkey("ctrl+alt+e", rodar_entrada_saida)

print("✅ Atalhos registrados:")
print("   CTRL + ALT + W para enviar WhatsApp")
print("   CTRL + ALT + P para manutenção preventiva")
print("   CTRL + ALT + A para automação")
print("   CTRL + ALT + E para Entrada e saida")
keyboard.wait()  # mantém rodando em background
