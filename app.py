
import openpyxl
import pandas as pd
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 


# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(30)

#Lê a planilha com pandas para filtrar contatos com status diferente de 'ok'
df = pd.read_excel('clientes.xlsx')
contatos_para_enviar = df[df['status'].astype(str).str.lower().str.strip() != 'ok']

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['clientes']

#Loop pelas linhas da planilha (a partir da 2ª linha)
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    status = linha[3].value

    #Pula se já está com status 'ok'
    if str(status).lower() == 'ok':
        continue

    # Verifica se há data de vencimento
    if vencimento:
        mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.link_do_pagamento.com'
    else:
        print(f'⚠️ Cliente {nome} está sem data de vencimento. Pulando...')
        continue

    #Envia mensagem via WhatsApp
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(5)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)

        #Atualiza o status para 'ok' na planilha
        linha[3].value = 'ok'  # Coluna "Status"

    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
            
    workbook.save('clientes.xlsx')
