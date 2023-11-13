import pandas as pd
# pip install python-docx
from docx import Document
import smtplib
from email.message import EmailMessage
from credenciais import meu_email, token


def ler_csv(arquivo: str):
    df = pd.read_csv(arquivo)
    return df

def criar_contrato(arquivo, lista):
    doc = Document(arquivo)
    for paragrafo in doc.paragraphs:
        for trechos in paragrafo.runs:
            for chave, valor in lista.items():
                if chave in trechos.text:
                    trechos.text = trechos.text.replace(chave, valor)
    doc.save(f"contratos/contrato de {lista['NOME']}.docx")
           

           
def enviar_email(arquivo, email, nome):
    msg = EmailMessage()
    msg['Subject'] = "Segue em anexo o contrato" # assunto
    msg['From'] = meu_email # remetente
    msg['To'] = email
    msg.set_content(f"Boa tarde {nome}, segue em anexo o seu contrato")
    msg.add_attachment(open(arquivo, "rb").read(), maintype="docx", subtype="docx", filename=f"Contrato de {nome}.docx ")

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        try:
            smtp.login(meu_email, token)
            smtp.send_message(msg)
            print(f"Email do {nome} Enviado com sucesso")

        except:
            print("Email não enviado")



dados = ler_csv("dados/dados.csv")
for key, items in dados.iterrows():
    criar_contrato("modelos/template.docx", {
        "NOME": items['Nome'],
        "ESTADO CIVIL": items['Estado Civil'],
        "PROFISSAO": items['Profissão'],
        "CPF1": items['CPF'],
        "RG1": items['RG'],
        "RUA": items['Endereço do Imovel'],
        "NUM": str(items['Nº da Rua']),
        "CIDADE": items['Cidade'],
        "ESTADO": items['Estado'],
        "VALOR": items['Valor da Locação'],
        "DATA": str(items['Data']).split(" ")[0] # lista = [data, hora]
    })
    enviar_email(f"contratos/contrato de {items['Nome']}.docx", items['Email'], items['Nome'])

print("\t\n Operação realizada com sucesso !!!\n")