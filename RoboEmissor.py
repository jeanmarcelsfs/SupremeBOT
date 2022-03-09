import pandas as pd
import pyperclip as ppc
import pyautogui as pag
import os
import time 
pag.PAUSE(0.5)

#Funcao Login no Sistema
def loginSistema(usuario, senha):
    try:
        pag.press("Win")
        pag.write("Infosis")
        time.sleep(3)
        pag.press("Enter")
        time.sleep(8)
        pag.write(usuario)
        pag.press("Enter")
        pag.write(senha)
        pag.press("Enter")
        time.sleep(5)
        return True
    except:
        print("ouve algum problema")
        return False
#Fim loginSistema    

#Funcao salvar PDF
def salvarPDF(nomeArquivo, diretorio):
    try:
        time.sleep(1)
        pag.write(nomeArquivo)
        pag.press("F4")
        time.sleep(1)
        pag.press("Esc")
        ppc.copy(diretorio)
        pag.hotkey("ctrl", "v")
        pag.press("enter")
        pag.hotkey("alt", "l")
        return True
    except:
        return False
#Fim salvar PDF

#Funcao Abre o Emissor de NF
def abreEmissor():
    pag.press("F11")
    time.sleep(4)
    pag.press("Esc")
    time.sleep(2)
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
        pag.press("b")
        pag.press("F12")
        pag.press("F8")
        pag.write(pedido)
        pag.press("Enter")
        pag.press("Enter")
        pag.press("Enter")
        time.sleep(3)
        pag.hotkey("alt", "g")
        pag.press("Enter")
        pag.hotkey("alt", "i")
        pag.press("Enter")
        pag.write("Boleto")
        pag.press("Enter")
        pag.press("Esc")
        pag.hotkey("Alt", "F")
       
        return True
    except:
        return False
#Fim Emite Boleto

#Funcao Emite NF
def emiteNF(pedido, obs, caminhoPDF):
    try:        
        pag.press("F10")
        pag.press("F8")
        pag.write(pedido)
        pag.press("Enter")
        pag.press("Enter")
        pag.press("Enter")
        pag.press("Esc")
        pag.hotkey("Alt", "m")
        pag.write(obs+" - ")
        pag.click(1130, 136)
        pag.press("Enter")
        pag.press("Enter")
        time.sleep(4)
        pag.hotkey("Alt", "i")
        pag.press("Enter")
        
        salvarPDF("Nota Fiscal", caminhoPDF)
        
        time.sleep(8)
        pag.press("Esc")
        pag.press("Esc")

        emiteBoleto(pedido)

        return True
    except:
        print("ouve algum problema")
        return False
#Fim emite NF

'''

I N I C I O

'''
loginSistema("Jean", "0409")
abreEmissor()

dataFrame = pd.read_excel(r'C:\Users\jean_\Desktop\_faturar.xlsx')
qtdeLinhas = dataFrame['Pedido'].count()
i = 0
while i < qtdeLinhas:    
    pedido = str(dataFrame.iloc[i][0])
    razaoSocial = str(dataFrame.iloc[i][6])
    diretorio = r"z:/Faturamento/Vendas/Frigorifico/TesteRobo/"
    nomePasta =  pedido + " - " + razaoSocial
    criaPasta(nomePasta, diretorio)    
    emiteNF(pedido, diretorio + nomePasta)
    i = i+1
#Fim While

