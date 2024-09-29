"""
a
"""
#   passo a passo de ler os dados das planilhas

#   ler planilha (nome, telefone e vencimento)
#   personalizar link do whatsapp, que automatiza a mensagem a ser enviada
#   esse link vai acelerar o processo de envio de informações

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(10)
pyautogui.hotkey('ctrl','w')

workbook = openpyxl.load_workbook('planilha.xlsx')
page = workbook['Sheet1']

for line in page.iter_rows(min_row=2):#min_row é referente a linha inicial da leitura
    #nome, telefone, mensagem
    nome = line[0].value
    telefone = line[1].value
    mensagem = line[2].value
    mensagem = f'{nome} {mensagem}'
    #mensagem a ser formatada: https://web.whatsapp.com/send?phone=&text=
    mensagemformatada = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(mensagemformatada)
    sleep(7)
    try:
        pyautogui.hotkey('enter')
        sleep(2)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'não foi possível mandar a mensagem para {nome}')

    
