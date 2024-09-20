import pyautogui
import time

print("Mova o mouse para a posição desejada...")
time.sleep(5)  # Tempo para você posicionar o mouse
position = pyautogui.position()
print("Posição do mouse:", position)