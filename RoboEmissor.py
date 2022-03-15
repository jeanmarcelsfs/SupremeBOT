import pandas as pd
import pyperclip as ppc
import pyautogui as pag
import os
import time 
from datetime import date

import ctypes  # An included library with Python install.
def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)



pag.PAUSE = 1

'''

R O B O   E M I S S O R

'''
#Funcao retorna Mes e Dia para criação de diretorio
def getMesDia():
    mes = time.strftime("%m")
    dia = time.strftime("%d")
    mesDia = ""
    if mes == "01":
        mesDia = mes + " - Janeiro/" + dia
    elif mes == "02":
        mesDia = mes + " - Fevereiro/" + dia
    elif mes == "03":
        mesDia = mes + " - Março/" + dia
    elif mes == "04":
        mesDia = mes + " - Abril/" + dia
    elif mes == "05":
        mesDia = mes + " - Maio/" + dia
    elif mes == "06":
        mesDia = mes + " - Junho/" + dia
    elif mes == "07":
        mesDia = mes + " - Julho/" + dia
    elif mes == "08":
        mesDia = mes + " - Agosto/" + dia
    elif mes == "09":
        mesDia = mes + " - Setembro/" + dia
    elif mes == "10":
        mesDia = mes + " - Outubro/" + dia
    elif mes == "11":
        mesDia = mes + " - Novembro/" + dia
    else:
        mesDia = mes + " - Dezembro/" + dia
    return mesDia
#Funcao Login no Sistema
def loginSistema(usuario, senha):
    try:
        pag.press("Win")
        pag.write("Infosis")
        time.sleep(2)
        pag.press("Enter")
        time.sleep(8)
        pag.write(usuario)
        pag.press("Enter")        
        pag.write(senha)
        pag.press("Enter")
        time.sleep(3)
        return True
    except:
        Mbox('E R R O', "Erro no Login", 0)
        return False
#Fim loginSistema    

#Funcao salvar PDF
def salvarPDF(nomeArquivo, diretorio):
    try:
        time.sleep(2)
        pag.write(nomeArquivo)
        pag.press("F4")
        time.sleep(2)
        pag.press("Esc")
        ppc.copy(diretorio)
        pag.hotkey("ctrl", "v")
        pag.press("enter")
        pag.hotkey("alt", "l")
        time.sleep(3)
        return True
    except:
        return False
#Fim salvar PDF

#Funcao Abre o Emissor de NF
def abreEmissor():
    pag.press("F11")
    time.sleep(7)
    pag.press("Esc")
    time.sleep(3)
#Fim Abre o Emissor

#Funcao Criar pasta no caminho especificado
def criaPasta(nomePasta, caminhoPasta):
    if not os.path.exists(caminhoPasta + nomePasta):
        os.makedirs(caminhoPasta+nomePasta)
#Fim criaPasta

#Funcao Emite Boletos
def emiteBoleto(pedido):
    try:
        pag.hotkey("alt", "o")
        time.sleep(2)
        pag.press("b")
        time.sleep(4)
        pag.press("F12")
        pag.press("F8")
        pag.write(pedido)
        pag.press("Enter")
        time.sleep(1)
        pag.press("Enter")
        pag.press("Enter")
        time.sleep(4)
        pag.hotkey("alt", "g")
        time.sleep(5)
        pag.press("Enter")
        time.sleep(5)
        pag.hotkey("alt", "i")
        pag.press("Enter")
        pag.write("Boleto")
        pag.press("Enter")
        pag.press("Esc")
        pag.hotkey("Alt", "F")
       
        return True
    except:
        Mbox('E R R O', "Erro no Emite Boleto", 0)
        return False
#Fim Emite Boleto

#Funcao Emite NF
def emiteNF(pedido, obs, caminhoPDF, codCli):
    try:        
        pag.press("F10")
        pag.press("F8")
        pag.write(pedido)
        pag.press("Enter")
        pag.press("Enter")
        pag.press("Enter")
        time.sleep(3)
        pag.press("Esc")
        pag.hotkey("Alt", "m")
        pag.write(obs+" - ")
        pag.click(1130, 136) #//Jean
        # pag.click(1420, 294) # Everton
        time.sleep(2)
        pag.press("Enter")
        pag.press("Enter")
        time.sleep(10)
        pag.hotkey("Alt", "i")
        pag.press("Enter")

        salvarPDF("Nota Fiscal", caminhoPDF)
                
        pag.press("Esc")
        time.sleep(8)
        pag.press("Esc")

        '''
        verifica quais clientes não serão impressos os boletos
        1817 ~ 1825: Mambo
        1124, 1677, 1665, 1488, 1455: Eistein

        '''
        listaCli = [1817, 1818, 1819, 1820, 1821, 1822, 1823, 1824, 1825, #Mambo
        1124, 1677, 1665, 1488, 1455 # Eistein
        ]

        if int(codCli) not in listaCli:
            emiteBoleto(pedido)

        return True
    except:
        Mbox("E R R O", "Erro ao Emitir NF", 0)
        return False
#Fim emite NF

'''

I N I C I O

'''
#loginSistema("Jean", "0409")
#abreEmissor()

dataFrame = pd.read_excel(r'faturar.xlsx')
qtdeLinhas = dataFrame['Pedido'].count()
i = 0
mesDia = getMesDia()

diretorio = r"z:/Faturamento/Vendas/Frigorifico/2022/" + mesDia  #Jean
#diretorio = r"E:/Faturamento/Vendas/Frigorifico/2022/" + mesDia #Servidor

if not os.path.exists(diretorio):
        os.makedirs(diretorio)

diretorio = diretorio+"/"

while i < qtdeLinhas:    
    pedido = str(dataFrame.iloc[i][0])
    razaoSocial = str(dataFrame.iloc[i][6])
    obs = str(dataFrame.iloc[i][8])
    codCli = str(dataFrame.iloc[i][5])    
    nomePasta =  pedido + " - " + codCli + " - " + razaoSocial
    criaPasta(nomePasta, diretorio)    
    #emiteNF(pedido, obs, diretorio + nomePasta, codCli)
    i = i+1
#Fim While
