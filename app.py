import openpyxl
import pandas as pd
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

# Abrir o WhatsApp Web no navegador
webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Carregando a planilha Excel
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['clientes']

# Loop principal para enviar mensagens um por um
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    status = linha[3].value


    # Pula se já está com status 'ok'(para nao mandar repetido)
    if status == 'ok':
        continue

    # Verifica se há data de vencimento e  Cria a mensagem personalizada
    if vencimento:
        mensagem = (
            f'Olá {nome} tudo bem ?, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. '
            'Chave pix: 4199999-9999'
        )
    else:
        print(f'⚠️ Cliente {nome} está sem data de vencimento. Pulando...')
        continue

      # Tenta enviar a mensagem via WhatsApp e cria link do whatsapp
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)                 # Abre o link no navegador
        sleep(5)                                              
        pyautogui.press('enter')                                # Pressiona Enter para enviar a mensagem
        sleep(5)                                               
        pyautogui.hotkey('ctrl', 'w')                           # Fecha a aba do WhatsApp
        sleep(5)                                                

        # Atualiza o status na planilha para apenas 'ok'
        linha[3].value = "ok"
        
    #Se tuver algum erro de envio ira para a planilha de erros
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv','a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}\n')
            


# Salva planilha
workbook.save('clientes.xlsx')
