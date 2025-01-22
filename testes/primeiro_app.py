import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from time import sleep

def is_valid_contact(contact):
    """
    Verifica se o texto fornecido é um número de contato válido no formato brasileiro.
    """
    # Padrão para números de telefone brasileiros: (XX) XXXXX-XXXX ou (XX) XXXX-XXXX
    pattern = r"^\(?\d{2}\)?\s?\d{4,5}-\d{4}$"
    return bool(re.match(pattern, contact))

# Carregar a planilha existente
planilha = openpyxl.load_workbook('dados.xlsx')
dados = planilha['Planilha1']

driver = webdriver.Chrome()
# Entrar no maps
driver.get("https://www.google.com.br/maps/preview")
sleep(5)
# Escrever e buscar 
escreve = driver.find_element(By.XPATH, '//*[@id="searchboxinput"]')
sleep(2)
escreve.send_keys("Academia em gramados")
sleep(1)
botao_pesquisa = driver.find_element(By.XPATH, '//*[@id="searchbox-searchbutton"]/span')
sleep(1)
botao_pesquisa.click()
sleep(3)
# Filtrar 4.0
filtro = driver.find_element(By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div[1]/div/div[2]/div[2]/div[1]/button/div/div[1]')
filtro.click()
sleep(1)
estrelas = driver.find_element(By.XPATH, '//*[@id="action-menu"]/div[6]')
estrelas.click()
sleep(3)

# Adicionar a lógica de rolar até o final da página
last_height = driver.execute_script("return document.body.scrollHeight")

# Inicializar a quantidade de elementos
quantidade_anterior = 0

while True:
    # Rola até o final da página
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    sleep(2)
    
    # Verifica a quantidade de elementos
    resultados = driver.find_elements(By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div')
    quantidade_atual = len(resultados)

    print(f"Quantidade atual de resultados: {quantidade_atual}")

    # Se a quantidade atual for maior que a anterior, clica no último elemento
    if quantidade_atual > quantidade_anterior:
        ultimo_elemento = resultados[-1]
        driver.execute_script("arguments[0].scrollIntoView();", ultimo_elemento)
        sleep(1)

        # Espera explícita para garantir que o elemento esteja clicável
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div[last()]')))
        sleep(1)
        try:
            ultimo_elemento.click()
        except Exception as e:
            print(f"Erro ao clicar no último elemento: {e}")
        
        quantidade_anterior = quantidade_atual

    # Verifica se a quantidade de elementos não mudou desde a última verificação
    sleep(2)
    novos_resultados = driver.find_elements(By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[1]/div[1]/div')
    nova_quantidade = len(novos_resultados)

    if nova_quantidade == quantidade_atual:
        break

# Descobre quantos elementos existem no XPath
elements = driver.find_elements(By.XPATH, '//a[@class="hfpxzc"]')  # Captura todos os elementos com o XPath desejado
total_elements = len(elements)  # Conta o número de elementos

# Loop para clicar em cada elemento
for i in range(1, total_elements + 1):  # Índices no XPath começam em 1
    try:
        # Gera o XPath dinâmico com o índice
        xpath = f'(//a[@class="hfpxzc"])[{i}]'
        
        # Localiza e clica no elemento
        element = driver.find_element(By.XPATH, xpath)
        element.click()
        sleep(2)  # Espera 2 segundos após clicar

        # Encontrar e copiar os campos "nome" e "contato" do primeiro resultado
        extrair_nome = driver.find_element(By.XPATH, '//h1[text()]')
        nomes = extrair_nome.text

        contato = "Não tem contato"
        for j in range(2, 6):  # Tentativa de arr[2] até arr[5]
            try:
                extrair_contato = driver.find_element(By.XPATH, f'(//div[@class="Io6YTe fontBodyMedium kR99db fdkmkc "])[{j}]')
                texto_contato = extrair_contato.text
                if is_valid_contact(texto_contato):
                    contato = texto_contato
                    break
            except:
                continue

        print(f"Nome: {nomes}, Contato: {contato}")

        # Adicionar os dados à planilha
        dados.append([nomes, contato])

        # Salvar a planilha
        planilha.save('dados.xlsx')

        sleep(5)
        
        # Clica no botão para voltar
        anterior = driver.find_element(By.XPATH, '//*[@id="omnibox-singlebox"]/div/div[1]/button/span')
        anterior.click()
        sleep(3)  # Espera 3 segundos antes de clicar no próximo elemento
    except Exception as e:
        print(f"Erro ao clicar no elemento {i}: {e}")
        continue

input("Pressione Enter para fechar...")

# Fecha o navegador
driver.quit()
