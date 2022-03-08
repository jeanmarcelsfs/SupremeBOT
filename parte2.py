
import pyautogui as pag
import pyperclip as ppc
import time
pag.PAUSE = 0.5     


#time.sleep(5)
#pag.hotkey("alt", "Tab", "Tab", "Tab") #quando for exe apagar 1
pag.click(967, 746)
pag.press("F12")
time.sleep(1)
pag.write("faturar")
pag.press("F4")
time.sleep(1)
pag.press("Esc")
ppc.copy(r"%homepath%\desktop")
pag.hotkey("ctrl", "v")
''''
pag.press("enter")
pag.press("tab")
pag.press("tab")
pag.press("tab")
pag.press("tab")
pag.press("tab")
pag.press("tab")

pag.press("enter")
pag.hotkey("alt" "F4")
'''