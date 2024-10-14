from PIL import ImageGrab
import pyautogui as pag
from time import sleep
import pygetwindow as gw
import pyperclip
import re

def ativar_janela(titulo):
    try:
        janela = gw.getWindowsWithTitle(titulo)[0]
        janela.activate()
        sleep(0.5)
    except IndexError:
        print(f"Janela '{titulo}' não encontrada.")

def automacao():
    x_excel, y_excel = 235, 327
    x_winthor_campo, y_winthor_campo = 659, 75
    x_botao_procurar, y_botao_procurar = 664, 107
    x_copiar_texto, y_copiar_texto = 1016, 452
    x_excel_coluna_valor, y_excel_coluna_valor = 977, 325

    linhas = 10  

    for i in range(linhas):

        ativar_janela("Excel")
        pag.moveTo(x_excel, y_excel + (i * 20), duration=4)  
        pag.click()
        pag.hotkey('ctrl', 'c') 

        sleep(4) 

        ativar_janela("2013 - Pesquisa de embalagens (Mata Burro)")
        pag.moveTo(x_winthor_campo, y_winthor_campo, duration=4)
        pag.click()
        pag.hotkey('ctrl', 'v')  

        sleep(4) 

        
        pag.moveTo(x_botao_procurar, y_botao_procurar, duration=4)
        pag.click()

        sleep(4) 

        
        pag.moveTo(x_copiar_texto, y_copiar_texto, duration=4)
        pag.doubleClick()

        sleep(4)
        pag.hotkey('ctrl', 'c') 

        sleep(4) 
      
        texto_capturado = pyperclip.paste()
    
        preco_venda = re.search(r'P\.venda\s+([\d,]+)', texto_capturado)
        if preco_venda:
            preco_venda = preco_venda.group(1)  
            
            ativar_janela("Excel")
            pag.moveTo(x_excel_coluna_valor, y_excel_coluna_valor + (i * 20), duration=4)
            pag.click()
            pag.click()
            pag.write(preco_venda) 

        sleep(4)

if __name__ == "__main__":
    automacao()
