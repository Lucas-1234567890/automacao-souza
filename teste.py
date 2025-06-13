import pyautogui
import time

print("⚠️ Coloca o mouse onde você quer capturar... (você tem 5 segundos)")
time.sleep(5)

x, y = pyautogui.position()
print(f"📍 Posição capturada: x={x}, y={y}")
