from botcity.core import DesktopBot
from botcity.plugins.email import BotEmailPlugin
from botcity.maestro import *
import pandas as pd
import PyPDF2

BotMaestroSDK.RAISE_NOT_CONNECTED = False

def ler_pdf(caminho_pdf):
    texto_pdf = []
    try:
        with open(caminho_pdf, 'rb') as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)
            num_paginas = len(leitor.pages)
            for i in range(num_paginas):
                pagina = leitor.pages[i]
                texto = pagina.extract_text()
                texto_pdf.append(texto)
    except Exception as e:
        print(f"Erro ao ler o PDF: {e}")
    return texto_pdf

def analisar_texto(texto_pdf):
    resultado = []
    for texto in texto_pdf:
        linhas = texto.strip().split('\n')
        linhas = linhas[1:]
        print(linhas)
        for linha in linhas:
            partes = linha.split()
            if len(partes) < 3:
                print(f"Linha com formato inesperado: {linha}")
                continue
            cpf = partes[-2]
            status = partes[-1]
            nome = ' '.join(partes[:-2])

            if status == 'Não':
                resultado.append({'Nome': nome, 'CPF': cpf, 'Status': status})

    return resultado

def salvar_em_excel(dados, caminho_excel):
    try:
        df = pd.DataFrame(dados)
        df.to_excel(caminho_excel, index=False)
        print(f"Dados salvos em {caminho_excel}")
    except Exception as e:
        print(f"Erro ao salvar em Excel: {e}")

def processar_pdf(caminho_pdf, caminho_excel):
    textos = ler_pdf(caminho_pdf)
    dados = analisar_texto(textos)
    salvar_em_excel(dados, caminho_excel)
    

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
