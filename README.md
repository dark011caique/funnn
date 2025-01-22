# Nome do Projeto

## Configuração do Ambiente Virtual
    python -m venv venv
    venv\Scripts\activate
    pip install -r requirements.txt

### Windows

1. Abra o terminal (Prompt de Comando ou PowerShell).
2. Navegue até o diretório do seu projeto.
3. Crie o ambiente virtual com o comando:
   ```bash
   python -m venv venv

## Abrir o Google com seu perfil
    "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebug"


### Explicação:
- O `README.md` contém instruções específicas para configurar o ambiente virtual em ambos os sistemas operacionais (Windows e Linux).
- A seção para rodar o Google Chrome com depuração remota está incluída com o comando apropriado para cada sistema operacional.
- O arquivo `requirements.txt` é usado para garantir que todas as dependências do projeto sejam instaladas de forma consistente.


### Ordem de execução para rodar o codigo:
    ### Passo 1: cmd
        "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebug"

    ### Passo 2: entra no ambiente virtual e roda o codigo
         venv\Scripts\activate
         app.py
