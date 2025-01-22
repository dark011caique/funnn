from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Caminho para o diretório do perfil do Chrome (onde seu login está)
chrome_user_data_dir = r"C:\Users\Win10\AppData\Local\Google\Chrome\User Data"# Substitua pelo caminho correto

# Configuração do Selenium para se conectar ao navegador já aberto
options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "localhost:9222")  # Conecta à porta de depuração
options.add_argument(f"user-data-dir={chrome_user_data_dir}")  # Usar o diretório de dados do usuário

# Usando o WebDriver Manager para baixar o ChromeDriver automaticamente
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Abra o WhatsApp Web
driver.get("https://web.whatsapp.com/")
print("Conexão estabelecida com o Chrome no modo de depuração.")

# Aguarde o tempo necessário para a página carregar
sleep(10)  # Aguarde o tempo necessário para carregar a página e os cookies

# Número de destino e mensagem
phone_number = "+5511949266294"  # No formato internacional
message = "Olá! Esta é uma mensagem enviada com Selenium."

# Construa a URL do número no WhatsApp Web
url = f"https://web.whatsapp.com/send?phone={phone_number}&text={message}"

# Abra a conversa com o número
driver.get(url)

# Aguarde a página carregar completamente
sleep(15)

# Enviar a mensagem
send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
send_button.click()

# Aguarde um tempo para garantir que a mensagem foi enviada
sleep(10)

print(f"Mensagem enviada para {phone_number}.")

# Feche o navegador
driver.quit()
