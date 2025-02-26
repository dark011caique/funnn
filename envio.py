from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import tkinter as tk
from openpyxl import load_workbook  # Certifique-se de que o openpyxl está instalado

def create_interface():
    def enviar_e_sair():
        global descricao
        descricao = campo_texto.get("1.0", tk.END).strip()
        root.destroy()

    root = tk.Tk()
    root.title("Enviar Mensagem")

    # Layout da interface gráfica
    label = tk.Label(root, text="Digite a mensagem:")
    label.pack()

    campo_texto = tk.Text(root, width=50, height=10)  # Aumentei o tamanho do campo de texto
    campo_texto.pack()

    botao_enviar = tk.Button(root, text="Enviar", command=enviar_e_sair)
    botao_enviar.pack()

    root.mainloop()

    return descricao


def enviar_mensagens_whatsapp(file_path, chrome_user_data_dir):
    # Configuração do Selenium para se conectar ao navegador já aberto
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "localhost:9222")  # Porta de depuração do Chrome
    options.add_argument(f"user-data-dir={chrome_user_data_dir}")  # Diretório de perfil do usuário

    # Inicia o WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Acessa o WhatsApp Web
    driver.get("https://web.whatsapp.com/")
    print("Aguardando carregamento do WhatsApp Web...")
    sleep(20)  # Aguarde o carregamento completo do WhatsApp Web (ajuste se necessário)

    # Carregar a planilha Excel
    try:
        workbook = load_workbook(file_path)
        sheet = workbook.active
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return

    # Obtém a mensagem digitada na interface gráfica
    conteudo = create_interface()

    # Loop pelos contatos na planilha
    for row in sheet.iter_rows(min_row=3, values_only=True):  # Começa na linha 3 para ignorar cabeçalhos
        nome = row[0]  # Nome na primeira coluna
        contato = str(row[1])  # Número de telefone na segunda coluna

        # Se o contato estiver vazio ou for "(Não tem contato)", pula para o próximo
        if not contato or contato == "(Não tem contato)":
            print(f"Ignorando {nome}, pois o contato não é válido.")
            continue

        mensagem = f"Olá {nome}! {conteudo}"

        # Construir a URL do WhatsApp Web para o contato
        url = f"https://web.whatsapp.com/send?phone={contato}&text={mensagem}"
        driver.get(url)

        print(f"Enviando mensagem para {nome} ({contato})...")
        sleep(10)  # Tempo para a página carregar completamente

        try:
            # Localiza e clica no botão de envio da mensagem
            send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
            send_button.click()
            print(f"Mensagem enviada para {nome} ({contato}).")
        except Exception as e:
            print(f"Falha ao enviar mensagem para {nome} ({contato}). Erro: {e}")

        sleep(5)  # Tempo de espera antes de passar para o próximo contato

    print("Processo finalizado. Todas as mensagens foram enviadas.")
    driver.quit()


# Exemplo de uso
file_path = "caminho/para/sua/planilha.xlsx"  # Substitua pelo caminho correto do arquivo
chrome_user_data_dir = "caminho/para/seu/perfil"  # Substitua pelo caminho correto do perfil do Chrome

enviar_mensagens_whatsapp(file_path, chrome_user_data_dir)
