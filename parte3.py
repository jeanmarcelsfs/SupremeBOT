import pandas as pd
import pyperclip as ppc
import pyautogui as pag
import os
dataFrame = pd.read_excel(r'C:\Users\jean_\Desktop\_faturar.xlsx')


#print(dataFrame)
qtde = dataFrame['Pedido'].count()
i = 0
while i < qtde:
    print(dataFrame.iloc[i][0])
    caminho = dataFrame.iloc[i][0]
    caminho = caminho + " - " + dataFrame.iloc[i][6]
    if not os.path.exists(r'z:/TI/Teste/'+caminho):
        os.makedirs(r'z:/TI/Teste'+caminho)
    i = i+1
print("saiu do laço")


