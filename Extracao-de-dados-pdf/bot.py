from webdriver_manager.chrome import ChromeDriverManager
from botcity.plugins.email import BotEmailPlugin
from botcity.maestro import *
import pandas as pd
import os
from dotenv import load_dotenv
import pandas as pd
import PyPDF2


# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def enviar_email(user_email, user_senha, to_email, assunto, conteudo, arquivo_path ):
    
    email = BotEmailPlugin()

    email.configure_imap("imap.gmail.com", 993)

    email.configure_smtp("smtp.gmail.com", 587)

    email.login(user_email, user_senha)

    email.send_message(assunto, conteudo, to_email, attachments=arquivo_path, use_html=True)
     
    email.disconnect()

    print("E-mail enviado com sucesso!")



def parametro_emails():

    load_dotenv()

    to = ["sabrina.frazao@ifam.edu.br", "sabrinadasilvafrazao@gmail.com"]
    subject = "Relatorio_SUS"
    body = '''Bom dia! 
            Segue em anexo o relatorio das pessoas que não possuem número SUS'''
    files = ["docs\RelatorioSUS.xlsx"]


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

    print(f"ID da Tarefa é: {execucao.task_id}")
    print(f"Parâmetros da Tarefa são: {execucao.parameters}")

    bot = DesktopBot()
    bot.browse("http://www.botcity.dev")

    caminho_pdf = 'Controle_SUS.pdf'
    caminho_excel = 'Dados.xlsx'
    processar_pdf(caminho_pdf, caminho_excel)
    
def not_found(label):
    print(f"Element not found: {label}")
    
    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = ChromeDriverManager().install()

    # Opens the BotCity website.
    #bot.browse("https://www.botcity.dev")

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
