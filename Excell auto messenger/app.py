

#   passo a passo de ler os dados das planilhas

#   ler planilha (nome, telefone e mensagem)
#   personalizar link do whatsapp, que automatiza a mensagem a ser enviada
#   esse link vai acelerar o processo de envio de informações

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

# Primeira vez que abre o whatsapp é para que os próximos carregamentos de mensagens sejam mais rápidos
webbrowser.open('https://web.whatsapp.com/')
sleep(10)
pyautogui.hotkey('ctrl','w')

# Abertura do worksheet e armazenamento para futura leitura
workbook = openpyxl.load_workbook('planilha.xlsx')
page = workbook['Sheet1']
erros = []

for line in page.iter_rows(min_row=2):#min_row é referente a linha inicial da leitura
    #nome, telefone, mensagem
    nome = line[0].value
    telefone = line[1].value
    mensagem = line[2].value
    mensagem = f'{nome} {mensagem}' #formata para primeiro o nome e a mensagem a ser passada
    #mensagem a ser formatada: https://web.whatsapp.com/send?phone=&text=
    mensagemformatada = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    try:
        # utilizando o quotes, da biblioteca urllib.parse, é possível formatar a mensagem para o padrão do whatsapp web link
        webbrowser.open(mensagemformatada) # aqui é aberto o link formatado com a mensagem pronta para ser enviada
        sleep(10)
        pyautogui.hotkey('enter')
        sleep(2)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'não foi possível mandar a mensagem para {nome}') 
        #caso uma mensagem não seja passada para alguem, vai ficar salvo no terminal
        erros.append([nome,telefone])

#percebi que o código está rodando infinitamente, ainda estou verificando como parar ao final da lista

    
