import pandas as pd
import pyautogui
import openpyxl
import time
import tkinter as tk
from tkinter import messagebox

pyautogui.PAUSE = 1
time.sleep(2)
# diretorio = 'G:\\LORENA\\alterar preços\\alterar preços\\BaseAltPreco.xlsx'
# tabela_base = pd.read_excel(diretorio)
imagem_erro = 'fotos/errototvs.png'

workbook = openpyxl.load_workbook('BaseAltPreco.xlsx', data_only=True)  
sheet_produtos = workbook['Base']


def alerta_concluido():    
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    tk.messagebox.showinfo("Operação Concluída", "O bot concluiu a operação com sucesso!")
    root.destroy()  # Fecha a janela de alerta

def alerta_erro():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    tk.messagebox.showinfo("Operação Falhou", "TOTVS excedeu o numero de licenças. Tente novamente.")
    root.destroy()  # Fecha a janela de alerta


for linha in sheet_produtos.iter_rows(min_row=2):
        #Escreve o código
            codigo = str(linha[2].value)            
            pyautogui.write(codigo)

            #passa para próximo campo   
            pyautogui.press('tab')

            #escreve novo valor 
            novo_preco = str("{:.2f}".format(linha[3].value).replace('.', ','))
            pyautogui.write(novo_preco)
 

            #atualiza
            pyautogui.press('enter')
            linha[8].value = 'ALTERADO'
            workbook.save('BaseAltPreco.xlsx')
        
            time.sleep(3)
            pyautogui.press('enter')
            pyautogui.press('tab')
           
alerta_concluido()










    