import PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

# Configurar o tema do PySimpleGUI
sg.theme('Reddit')

# Definir o layout da janela
tela_mensagem = [
    [sg.Text('Descrição')],
    [sg.Input(key='descricao', size=(20, 5))],  # Ajustando o tamanho do campo
    [sg.Button('Enviar')]
]

# Criar a janela
janela = sg.Window('Enviar', layout=tela_mensagem)

# Loop de eventos
while True:
    event, values = janela.read()
    if event == sg.WIN_CLOSED or event == 'Enviar':
        break

# Fechar a janela
janela.close()

# Pegar os valores de entrada
descricao = values['descricao']  # Corrigido: 'descricao' em vez de 'descriacao'

# Concatenar os valores e buscar no Google Maps
enviar = f'{descricao}'


# Iniciar o driver do navegador (certifique-se de que o driver do navegador está no PATH)
driver = webdriver.Chrome()

# Entrar no Google Maps
driver.get("https://www.google.com.br/maps/preview")
sleep(5)

# Escrever e buscar 
escreve = driver.find_element(By.XPATH, '//*[@id="searchboxinput"]')
sleep(2)
escreve.send_keys(enviar)
sleep(1)

# Pressionar Enter para buscar
escreve.send_keys(u'\ue007')

input('precione enter para fechar')