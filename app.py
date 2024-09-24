import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(30)  # Tempo para você escanear o QR Code

# Ler planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
pageClients = workbook['Planilha1']

for idx, line in enumerate(pageClients.iter_rows(min_row=2), start=2):
    # Armazenar dados: nome, telefone, CashBack
    name = line[0].value
    phone = line[1].value
    cashBack = line[2].value
    vencimento = pageClients['F2'].value

    # Verificação se todos os dados (nome, telefone, cashback) estão vazios
    if not name and not phone and not cashBack:
        print("Fim do envio. Todas as células estão vazias.")
        pageClients.cell(row=idx, column=4, value='Fim da execução')
        workbook.save('clientes.xlsx')
        break # Interrompe o loop quando todas as células estão vazias

    # Verificação básica se dados estão presentes
    if not name or not phone or not cashBack:
        print(f"Dados faltando para {name}. Pulando...")
        pageClients.cell(row=idx, column=4, value='Dados incompletos')
        continue

    # Mensagem a ser enviada
    message = f'Olá {name}, tudo bem? Polo Wear SP Market, passando para lembra-la que você tem um valor de desconto em seu cashback de R${cashBack} vinculado ao seu CPF,  venha resgatar, você só precisa fazer uma compra do dobro do valor do bônus 😃, o desconto será abatido no máximo de 50%, da sua compra, o mesmo expira em {vencimento.strftime('%d/%m/%y')}.'

    try:
        # Gerar link do WhatsApp com a mensagem
        linkMessageWhatsapp = f'https://web.whatsapp.com/send?phone={phone}&text={quote(message)}'
        webbrowser.open(linkMessageWhatsapp)
        sleep(10)  # Tempo para carregar a janela do WhatsApp com o número
        
        # Tentar localizar a seta de envio da mensagem
        arrow = pyautogui.locateCenterOnScreen('arrow.png')
        
        if arrow:
            sleep(5)
            pyautogui.click(arrow[0], arrow[1])  # Clicar na seta para enviar
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')  # Fechar a aba após o envio
            pageClients.cell(row=idx, column=4, value='Sucesso ao enviar mensagem')
        else:
            print(f"Seta de envio não encontrada para {name}. Verifique a imagem 'arrow.png'.")
            pageClients.cell(row=idx, column=4, value='Seta de envio não encontrada')
    
    except Exception as e:
        print(f'Erro ao enviar mensagem para {name}')
        pageClients.cell(row=idx, column=4, value=f'Erro ao enviar mensagem')
    
try:
    workbook.save('clientes.xlsx')
    print("Arquivo salvo com sucesso.")
except Exception:
    print(f"Erro ao salvar arquivo")


