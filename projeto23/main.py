import pandas as pd
import pyautogui as pt
import os
import time 

data_venc = str(input("Digite a data de venc: "))
fornecedor = str(input("Fornecedor: "))

pt.FAILSAFE = True  # Mover para canto superior esquerdo
df = pd.read_excel("projeto23/Planilha_p.xlsx") 

# Conta o número de linhas
num_linhas = df.shape[0]

os.system('start notepad')  # Apenas para testes
time.sleep(0.5)

i = 0
while i < num_linhas: 
    valorID = df.loc[i, "ID"]
    valorVAL = df.loc[i, "Total Prestação"]
    valorDATA = pd.to_datetime(df.loc[i, "Emissão"]).strftime("%d/%m/%Y")
    valorPROP = df.loc[i, "PROP"]

    # Primeira parte do lançamento
    pt.PAUSE = 0.2
    pt.write(f'{valorID}')
    pt.press('tab')

    pt.write('1')
    pt.press('tab')

    pt.write('221')
    pt.press('tab')

    pt.write(f'{valorDATA}')
    pt.press('tab')
    pt.press('tab')

    pt.write(data_venc)  # Vencimento - data única
    pt.press('tab')

    for _ in range(5):
        pt.press('down')
    
    pt.press('tab')
    pt.press('tab')
    pt.press('down')
    pt.press('tab')
    
    for _ in range(2):
        pt.press('down')
    pt.press('tab')

    pt.write(f'{valorVAL}')
    pt.press('tab')
    pt.press('tab')
    pt.press('tab')

    # Salgueiro - Filial CE
    if fornecedor == "221":
        if valorPROP == "TRUE":
            apropriacao = "8001"
            conta = "101"
        else:
            apropriacao = "8000"
            conta = "501"
    # Serra Talhada - Matriz
    elif fornecedor == "38":
        if valorPROP == "TRUE":
            apropriacao = "5222"
            conta = "101"
        else:
            apropriacao = "5036"
            conta = "501"
    # Arcoverde - Filial PE
    elif fornecedor == "10750":  
        if valorPROP == "TRUE":
            apropriacao = "5222"
            conta = "101"
        else:
            apropriacao = "5036"
            conta = "501"
    else:
        print("Fornecedor não reconhecido")
        break 
    pt.write(apropriacao)
    pt.press('tab')

    pt.write(conta)
    pt.press('tab')

    pt.write('.')
    pt.press('tab')

    pt.press('enter')

    time.sleep(1.5)

    # Segunda parte do lançamento
    for _ in range(5):
        pt.press('tab')

    time.sleep(1.5)
    
    pt.press('tab')
    pt.press('enter')
    
    for _ in range(7):
        pt.press('tab')
    
    pt.press('enter')

    time.sleep(2)  # Esperar para não criar problema com aceleração

    # Nova linha será processada
    i += 1
