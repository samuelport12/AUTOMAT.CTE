import pandas as pd
import pyautogui as pt
import os
import time 

df = pd.read_excel("projeto23\Planilha.xlsx")
#print(df.head(1))
os.system('start notepad')
time.sleep(0.1)

a=0
while a < 5:
    valorID = df.loc[a,"ID"] #LOC[+LINHA,COLUNA]
    
    valorVAL = df.loc[a,"VALOR"]
    
    valorDATA = df.loc[a,"DATA"]

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

    pt.write('21/04/2025') # Vencimento - data unica
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

    pt.write("8000")
    pt.press('tab')

    pt.write('501')
    pt.press('tab')

    pt.write('.')
    pt.press('tab')

    pt.press('enter')

    a += 1

