import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

# Abrir o whats e esperar um pouco para que carregue
webbrowser.open('https://web.whatsapp.com/')
sleep(20)  # Ajuste de Delay em decorrencia da sua internete/maquina

# Carregando a planilha execel e selecionando em seguida a pagina de nome Sheet1
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

# Interação com os elementos da Sheet1 iniciado aparti da linha 2, evitando o cabeçalho
for linha in pagina_clientes.iter_rows(min_row=2, values_only=True):
    nome = linha[0]
    telefone = linha[1]
    mensagem = f'Olá {nome}. O seu contrato de amizade com o Senhor Pedro Luka, vulgo Xcorpion se encerrou hoje às 15:00.\nPara continuar a sua amizade visite o link a seguir: https://www.youtube.com/watch?v=-f-emAWnJGQ&list=RDCvlLKcY16qY&index=5'

    try:
        link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem)
        sleep(20)

        # Localizando o botão de enviar do whats e em seguida clicar
        seta = pyautogui.locateCenterOnScreen('seta1.png')
        if seta is not None:
            sleep(5)
            pyautogui.click(seta[0], seta[1])
            sleep(5)
        else:
            # Caso haja um erro de localização de imagem esse aviso irá nos ajudar
            raise Exception("\033[1; 31; 41mNão foi possivel localizar botão de enviar na tela")
        
        # fechar a aba aberta
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except Exception as e:
        print(f'Não foi possível enviar a mensagem para {nome}. Erro: {e}')
        # Criando um arquivo para sabermos melhor para quem não foi enviado a mensagem em caso de erro
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}\n')
