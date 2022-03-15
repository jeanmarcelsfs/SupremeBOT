import pyautogui as pag
import time

import ctypes  # An included library with Python install.
def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)



#time.sleep(3)
print(pag.position())
resposta = Mbox('meu Titulo', str(pag.position()), 0)
print(resposta)