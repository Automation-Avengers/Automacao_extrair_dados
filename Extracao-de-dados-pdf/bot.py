from botcity.web import WebBot, Browser, By

from webdriver_manager.chrome import ChromeDriverManager
from botcity.plugins.email import BotEmailPlugin
from botcity.maestro import *
import pandas as pd
import os
from dotenv import load_dotenv
import pandas as pd
import PyPDF2
from webdriver_manager.chrome import ChromeDriverManager


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
                        partes = linha.split()
                        if len(partes) >= 4 and partes[-1] == 'Não':  # Verifica se "Não" é a última palavra
                            nome = ' '.join(partes[:2])  # Tudo antes do CPF
                            cpf = partes[-2]  # CPF é o terceiro a partir do fim
                            tem_cartao = partes[-1]  # Sim ou Não
                            dados_extraidos.append([nome, cpf, tem_cartao])
                else:
                    print(f"Nenhum texto encontrado na página {page_num+1}.")
            return dados_extraidos
    except Exception as e:
        print(f"Erro ao extrair dados do PDF: {e}")
        return []

# Função para salvar dados em formato de tabela no Excel
def salvar_em_excel(dados, caminho_excel):
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

# Função principal para processar o PDF e salvar os dados no Excel
def processar_pdf(caminho_pdf, caminho_excel):
    dados = extrair_dados_pdf(caminho_pdf)
    if dados:
        salvar_em_excel(dados, caminho_excel)
        print("Dados processados e salvos com sucesso.")
    else:
        print("Nenhum dado foi extraído do PDF.")


# Envio do relátorio pelo email
def enviar_email(user_email, user_senha, to_email, assunto, conteudo, arquivo_path ):
    email = BotEmailPlugin()

    email.configure_imap("imap.gmail.com", 993)

    email.configure_smtp("smtp.gmail.com", 587)

    email.login(user_email, user_senha)

    email.send_message(assunto, conteudo, to_email, attachments=arquivo_path, use_html=True)
     
    email.disconnect()

    print("E-mail enviado com sucesso!")


# Parametro para o envio dos emails
def parametro_emails():

    load_dotenv()

    to = ["sabrina.frazao@ifam.edu.br", "sabrinadasilvafrazao@gmail.com", "jonassantos2302@gmail.com", "diego.eimec@gmail.com", "antonio.fernandes@icomp.ufam.edu.br"]
    subject = "Relatorio_SUS"
    body = '''Bom dia! 
            Segue em anexo o relatorio das pessoas que não possuem número SUS'''
    files = ["docs/relatorio_cartao_sus.xlsx"]


    enviar_email(
        user_email="sabrinadasilvafrazao@gmail.com",
        user_senha= os.getenv('SMTP_PASSWORD'),
        to_email=to,
        assunto=subject,
        conteudo=body,
        arquivo_path=files
    )
    


def main():
    maestro = BotMaestroSDK.from_sys_args()
    execucao = maestro.get_execution()

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = ChromeDriverManager().install()

    # Opens the BotCity website.
    #bot.browse("https://www.botcity.dev")

    caminho_pdf = r"docs\Controle_SUS.pdf"
    caminho_excel = r"docs\relatorio_cartao_sus.xlsx"
    
    processar_pdf(caminho_pdf, caminho_excel)

    parametro_emails()
  

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
