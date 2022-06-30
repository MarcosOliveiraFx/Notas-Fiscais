from dataclasses import dataclass
from h11 import Data
from imap_tools import MailBox, AND
import pandas as pd
import xmltodict
import datetime
import os
import fsspec

# pegar emails de um remetente para um destinatário
username = ""
password = ''
# lista de imaps: https://www.systoolsgroup.com/imap/
meu_email = MailBox('mail.mamoeiro.com.br').login(username, password)
# criterios: https://github.com/ikvk/imap_tools#search-criteria
lista_emails = meu_email.fetch(AND(from_="xml@guarany.com.br", subject="NFe")) 
#for email in lista_emails:
#    print(email.subject)
#    print(email.text)
# pegar emails com um anexo específico

date = datetime.date(
    year=2022,
    month=6,
    day=29
)

lista_emails = meu_email.fetch(AND(from_="xml@guarany.com.br", date={date}))
#lista_anexo = []

lista = []
informacoes_anexo = ''
for i, email in enumerate(lista_emails):
    if i >= 0:
        for anexo in email.attachments:
            if "nfe.xml" in anexo.filename:
                informacoes_anexo = anexo.payload
                lista.append(informacoes_anexo) 
                               
#print(len(lista))
#print("")
#print("")
#print("")
#print(lista)

dictionary0 = xmltodict.parse(lista[2])
dictionary1 = xmltodict.parse(lista[3])

cnpj0 = []
cnpj1 = []

razao0 = dictionary0['nfeProc']['NFe']['infNFe']['dest']['CNPJ']
cnpj0.append(razao0)
razao1 = dictionary1['nfeProc']['NFe']['infNFe']['dest']['CNPJ']
cnpj1.append(razao1)

print("")
print("")
print("")
print(cnpj0)
print("")
print(cnpj1)
print("")
print("")
print("")

notas_fiscal0 = []
notas_fiscal1 = []

for dici in dictionary0['nfeProc']['NFe']['infNFe']['det']:
    produto = {
        'descricao': dici['prod']['xProd'],
        'ean': str(dici['prod']['cEAN'].rjust(13, '0')),
        'quantidade': int(float(dici['prod']['qCom']))
    }
    notas_fiscal0.append(produto)
      
for dici in dictionary1['nfeProc']['NFe']['infNFe']['det']:
    produto = {
        'descricao': dici['prod']['xProd'],
        'ean': str(dici['prod']['cEAN'].rjust(13, '0')),
        'quantidade': int(float(dici['prod']['qCom']))
    }
    notas_fiscal1.append(produto)
   
lista_nf_df0 = pd.DataFrame.from_dict(notas_fiscal0)
lista_nf_df1 = pd.DataFrame.from_dict(notas_fiscal1)


if cnpj0 == ['']:
    print("WEB GLAMOUR")
    print(lista_nf_df0)
    lista_nf_df0.to_excel("NFe Web Glamour Tarde.xlsx")
if cnpj0 == ['']:
    print("ARMAZEM MAMOEIRO")
    print(lista_nf_df0)
    lista_nf_df0.to_excel("NFe Armazem Mamoeiro Tarde.xlsx")
if cnpj1 == ['']:
    print("WEB GLAMOUR")
    print(lista_nf_df1)
    lista_nf_df1.to_excel("NFe Web Glamour Tarde.xlsx")
if cnpj1 == ['']:
    print("ARMAZEM MAMOEIRO")
    print(lista_nf_df1)
    lista_nf_df1.to_excel("NFe Armazem Mamoeiro Tarde.xlsx")
    
