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
    
def enviar_email(user_email, user_senha, to_email, assunto, conteudo, arquivo_path ):
    email = BotEmailPlugin()
    email.configure_imap("imap.gmail.com", 993)
    email.configure_smtp("smtp.gmail.com", 587)
    email.login(user_email, user_senha)
    email.send_message(assunto, conteudo, to_email, attachments=arquivo_path, use_html=True)
    email.disconnect()

    print("E-mail enviado com sucesso!")
    
def parametro_emails():

    to = ["sabrina.frazao@ifam.edu.br", "sabrinadasilvafrazao@gmail.com"]
    subject = "Relatorio_SUS"
    body = '''Bom dia! 
            Segue em anexo o relatorio das pessoas que não possuem número SUS'''
    files = ["Dados.xlsx"]
    
    enviar_email(
    user_email="jonas.santos2302@gmail.com",
    user_senha="aaaaaaaaaaaaa",
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
    parametro_emails()
    
def not_found(label):
    print(f"Element not found: {label}")

if __name__ == '__main__':
    main()
