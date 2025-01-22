import PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

# Configurar o tema do PySimpleGUI
sg.theme('Reddit')

# Definir o layout da janela
tela_busca = [
    [sg.Text('Estabelecimentos comerciais')],
    [sg.Input(key='comercio')],
    [sg.Text('Cidade')],
    [sg.Input(key='local')],
    [sg.Button('Buscar')]
]

# Criar a janela
janela = sg.Window('Buscar', layout=tela_busca)

# Loop de eventos
while True:
    event, values = janela.read()
    if event == sg.WIN_CLOSED or event == 'Buscar':
        break

# Fechar a janela
janela.close()

# Pegar os valores de entrada
comercio = values['comercio']
local = values['local']

# Concatenar os valores e buscar no Google Maps
busca = f'{comercio} em {local}'

# Iniciar o driver do navegador (certifique-se de que o driver do navegador está no PATH)
driver = webdriver.Chrome()

# Entrar no Google Maps
driver.get("https://www.google.com.br/maps/preview")
sleep(5)

# Escrever e buscar 
escreve = driver.find_element(By.XPATH, '//*[@id="searchboxinput"]')
sleep(2)
escreve.send_keys(busca)
sleep(1)

# Pressionar Enter para buscar
escreve.send_keys(u'\ue007')

input('precione enter para fechar')