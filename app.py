import openpyxl
import pandas as pd
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import datetime

# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Carregar planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['clientes']

# Define tempo limite de 2 minutos em segundos
TEMPO_LIMITE_SEGUNDOS = 2 * 60  # 2 minutos

#Limpa "ok" antigos
for linha in pagina_clientes.iter_rows(min_row=2):
    status = linha[3].value
    if status and str(status).startswith('ok'):
        timestamp = float(status.split('|')[1])
        tempo_passado = datetime.datetime.now().timestamp() - timestamp
        if tempo_passado > TEMPO_LIMITE_SEGUNDOS:
            linha[3].value = "ok"  # limpa o ok

#Loop para envio de mensagens
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    status = linha[3].value

    # Pula se já está com status 'ok'
    if status and str(status).startswith('ok|'):
        continue

    # Verifica se há data de vencimento
    if vencimento:
        mensagem = f'Olá {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.link_do_pagamento.com'
    else:
        print(f'⚠️ Cliente {nome} está sem data de vencimento. Pulando...')
        continue

    # Envia mensagem via WhatsApp
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(5)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)

        # Atualiza o status para 'ok|timestamp'
        linha[3].value = f'ok|{datetime.datetime.now().timestamp()}'

    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')

# Salva planilha
workbook.save('clientes.xlsx')
