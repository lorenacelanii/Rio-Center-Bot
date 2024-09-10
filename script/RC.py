import pandas as pd
import pyautogui
import openpyxl
import time
import tkinter as tk
from tkinter import messagebox
import pyperclip

pyautogui.PAUSE = 1
time.sleep(2)

# Entrar na planilha
workbook = openpyxl.load_workbook('BaseAltPreco.xlsx', data_only=True)
sheet_produtos = workbook['Base']

def alerta_concluido():    
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    tk.messagebox.showinfo("Operação Concluída", "O bot concluiu a operação na RC")
    root.destroy()  # Fecha a janela de alerta

def alerta_erro():      
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    tk.messagebox.showinfo("Falha na operação", "Registro de alteração não concluído..tente novamente!")
    root.destroy()  # Fecha a janela de alerta


def erro_renovacao():
    pyautogui.press('enter')

time.sleep(2)

# Para cada linha na tabela
for linha in sheet_produtos.iter_rows(min_row=2):
    # Clica no campo para escrever código
    for _ in range(3):
        pyautogui.press('tab')  

    codigo = linha[2].value
    pyperclip.copy(codigo)
    pyautogui.hotkey('ctrl', 'v')

    # Pesquisa o código
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('v')

    # Pressiona Tab 12 vezes
    for _ in range(12):
        pyautogui.press('tab')

    # Copiar o valor selecionado
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.sleep(1)  # Aguardar um momento para garantir que o texto foi copiado

    # Obter o valor da área de transferência
    valor_copiado = pyperclip.paste().strip()

    # Verificar se o valor copiado é o mesmo que o valor do código
    if valor_copiado == codigo:
        print(f"Valor copiado ({valor_copiado}) é igual ao valor do código ({codigo}).")

        # Clicar em cancelar
        pyautogui.doubleClick(1212,150, duration=1)
        # OM - Manutenção de preços
        pyautogui.press('o')
        time.sleep(1)
        pyautogui.press('m')

        # Pausa para carregamento 
        time.sleep(7)

        # Coloca data de agendamento 
        pyautogui.doubleClick(432,273, duration=1)
        pyautogui.doubleClick(431,482, duration=1)

        # Insere novo preço
        for _ in range(5):
            pyautogui.press('tab')
        time.sleep(0.5)
        for _ in range(2):
            pyautogui.press('right')
        
        # Inserir Preço
        novo_preco = str("{:.2f}".format(linha[3].value).replace('.', ','))
        pyautogui.write(novo_preco)
        pyautogui.press('enter')

        # Clica em 'OK'
        pyautogui.press('tab')
        pyautogui.press('enter')

        # Remove data de agendamento
        for _ in range(7):
            pyautogui.press('tab')
        pyautogui.press('backspace')

        # Clica em confirmar 
        pyautogui.doubleClick(1312,122)
        time.sleep(3)
        
        localizacao = pyautogui.locateCenterOnScreen('fotos/registro_ok.png', confidence=0.8) 
        if localizacao is not None:
            print(f'codigo {codigo} foi alterado.')
            
            # Atualizar a planilha para marcar o preço como alterado
            linha[7].value = 'ALTERADO'
            # Salvar a planilha com as alterações
            workbook.save('BaseAltPreco.xlsx')

            # Clica em fechar 
            pyautogui.press('enter')
            time.sleep(3)
        else:
            print(f'codigo ({codigo}) não foi alterado')
            linha[7].value = 'NÃO ALTERADO'
            # Salvar a planilha com as alterações
            workbook.save('BaseAltPreco.xlsx')

        
    else:
        # Clicar em cancelar
        pyautogui.doubleClick(1212,150, duration=1)
        linha[7].value = 'CÓDIGO ERRADO'
        # Salvar a planilha com as alterações
        workbook.save('BaseAltPreco.xlsx')
        print(f"Valor copiado ({valor_copiado}) é diferente do valor do código ({codigo}).")

alerta_concluido()
    
