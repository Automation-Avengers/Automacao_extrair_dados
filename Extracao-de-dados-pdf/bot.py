"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/python-automations/web/
"""


# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *
from webdriver_manager.chrome import ChromeDriverManager

import PyPDF2

import pandas as pd

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


# Função para extrair dados do PDF
def extrair_dados_pdf(caminho_pdf):
    try:
        print("Iniciando a extração de dados do PDF...")
        with open(caminho_pdf, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            num_pages = len(reader.pages)
            print(f"PDF tem {num_pages} páginas.")
            dados_extraidos = []
            for page_num in range(num_pages):
                page = reader.pages[page_num]
                text = page.extract_text()
                if text:
                    linhas = text.split('\n')  # Dividir o texto por linhas
                    for linha in linhas:
                        # Verificar se a linha contém o formato esperado: Nome, CPF e "Sim"
                        partes = linha.split()
                        if len(partes) >= 4 and partes[-1] == 'Não':  # Verifica se "não" é a última palavra
                            nome = ' '.join(partes[:-3])  # Tudo antes do CPF
                            cpf = partes[-3]  # CPF é o terceiro a partir do fim
                            tem_cartao = partes[-1]  # Sim ou Não
                            dados_extraidos.append([nome, cpf, tem_cartao])
                else:
                    print(f"Nenhum texto encontrado na página {page_num+1}.")
            return dados_extraidos
    except Exception as e:
        print(f"Erro ao extrair dados do PDF: {e}")
        return []

# Função para salvar dados em formato de tabela no Excel
def salvar_dados_excel(dados, caminho_excel):
    try:
        if dados:
            print("Iniciando a criação do arquivo Excel...")
            # Criar DataFrame com as colunas Nome, CPF, Tem cartão do SUS
            df = pd.DataFrame(dados, columns=['Nome', 'CPF', 'Tem cartão do SUS'])
            df.to_excel(caminho_excel, index=False, sheet_name="Relatório Cartão SUS")  # Salvando como tabela
            print(f"Dados salvos no arquivo Excel: {caminho_excel}")
        else:
            print("Nenhum dado para salvar no Excel.")
    except Exception as e:
        print(f"Erro ao salvar dados no Excel: {e}")

            

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = ChromeDriverManager().install()

    # Opens the BotCity website.
    #bot.browse("https://www.botcity.dev")

    # Implement here your logic...

    caminho_pdf = 'Controle_SUS.pdf'  
    caminho_excel = 'relatorio_cartao_sus.xlsx'
 
    dados = extrair_dados_pdf(caminho_pdf)
    
    if dados:
     
        salvar_dados_excel(dados, caminho_excel)
          
    else:
        print("Nenhum dado foi extraído do PDF. O processo foi interrompido.")


    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )

def not_found(label):
    print(f"Element not found: {label}")
    

if __name__ == '__main__':

    
    main()
