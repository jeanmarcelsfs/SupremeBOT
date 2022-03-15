from time import strftime
import pandas as pd
import ctypes
from datetime import date

def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


mes = strftime("%m")
dia = strftime("%d")
mesf = mes
if mes == "03":
    mesf = mesf + " - Março/" + dia
print(mesf)   
