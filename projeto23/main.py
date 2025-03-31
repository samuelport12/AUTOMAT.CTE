import pandas as pd
import pyautogui as pt
import os
import time 

data_venc = str(input("digite a data de venc: "))
fornecedor = str(input("Fornecedor: ")

pt.FAILSAFE = True #MOVER PARA CANTO SUPERIOR ESQ.
df = pd.read_excel("projeto23\Planilha.xlsx")
# Conta o número de linhas
num_linhas = df.shape[0]

os.system('start notepad') #apenas para testes
time.sleep(0.1)

i = 0
while i <= num_linhas:
    valorID = df.loc[a,"ID"] #LOC[+LINHA,COLUNA]
    
    valorVAL = df.loc[a,"VALOR"]
    
    valorDATA = df.loc[a,"Emissão"]

    valorPROP = df.loc[a,"PROP"] 

    #primeira_parte_do_lançamento
    pt.PAUSE = 0.2   
    pt.write(f'{valorID}')  # Digita o ID
    pt.press('tab')

    pt.write('1')
    pt.press('tab')

    pt.write('221')
    pt.press('tab')

    pt.write(f'{valorDATA}')
    pt.press('tab')
    pt.press('tab')

    pt.write(data_venc) # Vencimento - data unica
    pt.press('tab')

    # # Pressiona seta para baixo 5 vezes
    for _ in range(5):
         pt.press('down')
    
    pt.press('tab')
    pt.press('tab')
    pt.press('down')  # 1 seta para baixo
    pt.press('tab')
    
    for _ in range(2):
         pt.press('down')
    pt.press('tab')

    pt.write(f'{valorVAL}')
    pt.press('tab')
    pt.press('tab')
    pt.press('tab')

#salgueiro - FILIAL CE 
    if fornecedor == "221":
        if valorPROP == "TRUE":
            apropriacao = "8001"
            conta = "101"
        else:
            apropriacao = "8000"
            conta = "501"
#Serra talhada - MATRIZ 
    elif fornecedor == "38":
        if valorPROP == "TRUE":
            apropriacao = "5222"
            conta = "101"
        else:
            apropriacao = "5036"
            conta = "501"
#arcoverde - FILIAL PE 
     elif fornecedor == "10750":
        if valorPROP == "TRUE":
            apropriacao = "5222"
            conta = "101"
        else:
            apropriacao = "5036"
            conta = "501"
    
    pt.write(apropriacao)
    pt.press('tab')

    pt.write(conta)
    pt.press('tab')

    pt.write('.')
    pt.press('tab')

    pt.press('enter')

    time.sleep(1.5) #esperar para não criar problema com aceleração
    #segunda_parte_do_lançamento 
    for _ in range(5):
    pt.press('tab')

    time.sleep(1.5)
    
    pt.press('tab')
    
    pt.press('enter')
    
    for _ in range(7):
        pt.press('tab')
    
    pt.press('enter')
    
    # Aguarda 2 segundos
    time.sleep(2) #esperar para não criar problema com aceleração

    #término_do_lançamento - nova linha será processada

    i += 1

