from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager

# Inicialize o driver do Selenium com o WebDriver Manager
driver = webdriver.Chrome()

# Abra o WhatsApp Web
driver.get("https://web.whatsapp.com/")

# Aguarde o tempo necessário para o QR Code ser escaneado (tempo depende do seu login)
print("Escaneie o QR Code!")
sleep(30)  # Aumente esse tempo se necessário

# Número de destino e mensagem
phone_number = "+5511949266294"  # No formato internacional
message = "Olá! Esta é uma mensagem enviada com Selenium."

# Construa a URL do número no WhatsApp Web
url = f"https://web.whatsapp.com/send?phone={phone_number}&text={message}"

# Abra a conversa com o número
driver.get(url)

# Aguarde a página carregar
sleep(15)

# Enviar a mensagem
send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
send_button.click()

# Aguarde um tempo para garantir que a mensagem foi enviada
sleep(10)

print(f"Mensagem enviada para {phone_number}.")

input("Pressione Enter para fechar...")
# Feche o navegador
driver.quit()
