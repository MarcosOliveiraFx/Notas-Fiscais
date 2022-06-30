from imap_tools import MailBox, AND
from dataclasses import dataclass
from h11 import Data
import pandas as pd
from attr import s
import xmltodict
import datetime
import requests
import fsspec
import os
import io

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



dictionary0 = xmltodict.parse(lista[0])
dictionary1 = xmltodict.parse(lista[1])
dictionary2 = xmltodict.parse(lista[2])
dictionary3 = xmltodict.parse(lista[3])



cnpj0 = []
cnpj1 = []
cnpj2 = []
cnpj3 = []
razao0 = dictionary0['nfeProc']['NFe']['infNFe']['dest']['CNPJ']
cnpj0.append(razao0)
razao1 = dictionary1['nfeProc']['NFe']['infNFe']['dest']['CNPJ']
cnpj1.append(razao1)
razao2 = dictionary2['nfeProc']['NFe']['infNFe']['dest']['CNPJ']
cnpj2.append(razao2)
razao3 = dictionary3['nfeProc']['NFe']['infNFe']['dest']['CNPJ']
cnpj3.append(razao3)

print("")
print("")
print("")
print(cnpj0)
print("")
print(cnpj1)
print("")
print(cnpj2)
print("")
print(cnpj3)
print("")
print("")
print("")

notas_fiscal0 = []
notas_fiscal1 = []
notas_fiscal2 = []
notas_fiscal3 = []



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
for dici in dictionary2['nfeProc']['NFe']['infNFe']['det']:
    produto = {
        'descricao': dici['prod']['xProd'],
        'ean': str(dici['prod']['cEAN'].rjust(13, '0')),
        'quantidade': int(float(dici['prod']['qCom']))
    }
    notas_fiscal2.append(produto)
for dici in dictionary3['nfeProc']['NFe']['infNFe']['det']:
    produto = {
        'descricao': dici['prod']['xProd'],
        'ean': str(dici['prod']['cEAN'].rjust(13, '0')),
        'quantidade': int(float(dici['prod']['qCom']))
    }
    notas_fiscal3.append(produto) 


  
lista_nf_df0 = pd.DataFrame.from_dict(notas_fiscal0)
lista_nf_df1 = pd.DataFrame.from_dict(notas_fiscal1)
lista_nf_df2 = pd.DataFrame.from_dict(notas_fiscal2)
lista_nf_df3 = pd.DataFrame.from_dict(notas_fiscal3)


lista_emails = meu_email.fetch(AND(from_="ti@mamoeiro.com.br", date={date}))
lista_Pedidos_Mamoeiro = []
informacoes_anexo_Pedidos_Mamoeiro = ''
for i, email in enumerate(lista_emails):
    if i >= 0:
        for anexo in email.attachments:
            if "Dag Mamoeiro.csv" in anexo.filename:
                informacoes_anexo_Pedidos_Mamoeiro = anexo.payload
                lista_Pedidos_Mamoeiro.append(informacoes_anexo_Pedidos_Mamoeiro)

lista_Pedido_Mamoeiro_df0 = pd.read_csv(io.StringIO(lista_Pedidos_Mamoeiro[0].decode('latin-1')),sep=";")
print("")
print("")
print("Pedidos Mamoeiro 09:00")
print(lista_Pedido_Mamoeiro_df0)

lista_Pedido_Mamoeiro_df1 = pd.read_csv(io.StringIO(lista_Pedidos_Mamoeiro[1].decode('latin-1')),sep=";")
print("")
print("")
print("Pedidos Mamoeiro 11:10")
print(lista_Pedido_Mamoeiro_df0)

lista_Pedido_Mamoeiro_df2 = pd.read_csv(io.StringIO(lista_Pedidos_Mamoeiro[2].decode('latin-1')),sep=";")
print("")
print("")
print("Pedidos Mamoeiro 13:30")
print(lista_Pedido_Mamoeiro_df2)

lista_emails = meu_email.fetch(AND(from_="ti@mamoeiro.com.br", date={date}))
lista_Pedidos_Webglamour = []
informacoes_anexo_Pedidos_Webglamour = ''
for i, email in enumerate(lista_emails):
    if i >= 0:
        for anexo in email.attachments:
            if "Dag Web Glamour.csv" in anexo.filename:
                informacoes_anexo_Pedidos_Webglamour = anexo.payload
                lista_Pedidos_Webglamour.append(informacoes_anexo_Pedidos_Webglamour)
                
lista_Pedido_Webglamour_df0 = pd.read_csv(io.StringIO(lista_Pedidos_Webglamour[0].decode('latin-1')),sep=";")
print("")
print("")
print("Pedidos Webglamour 09:00")
print(lista_Pedido_Webglamour_df0)

lista_Pedido_Webglamour_df1 = pd.read_csv(io.StringIO(lista_Pedidos_Webglamour[1].decode('latin-1')),sep=";")
print("")
print("")
print("Pedidos Webglamour 11:10")
print(lista_Pedido_Webglamour_df1)

lista_Pedido_Webglamour_df2 = pd.read_csv(io.StringIO(lista_Pedidos_Webglamour[2].decode('latin-1')),sep=";")
print("")
print("")
print("Pedidos Webglamour 13:30")
print(lista_Pedido_Webglamour_df2)


todos_pedidos_mamoeiro = pd.concat((lista_Pedido_Mamoeiro_df0, lista_Pedido_Mamoeiro_df1, lista_Pedido_Mamoeiro_df2), ignore_index=True)
print("")
print("")
print("Todos Pedidos da Mamoeiro")
print(todos_pedidos_mamoeiro)
todos_pedidos_mamoeiro.to_excel("todos_pedidos_mamoeiro.xlsx")

todos_pedidos_webglamour = pd.concat((lista_Pedido_Webglamour_df0, lista_Pedido_Webglamour_df1, lista_Pedido_Webglamour_df2), ignore_index=True)
print("")
print("")
print("Todos Pedidos da Web Glamour")
print(todos_pedidos_webglamour)
todos_pedidos_webglamour.to_excel("todos_pedidos_web_glamour.xlsx")


if cnpj0 == ['']:
    if cnpj0 == cnpj2:
        print("Todas Notas Fiscais Web Glamour")
        todas_notas_webglamour = pd.concat((lista_nf_df0, lista_nf_df2), ignore_index=True)
        todas_notas_webglamour.to_excel("Todas_Notas_Web_Glamour.xlsx")
        print(todas_notas_webglamour)
        
        
if cnpj0 ==  ['']:
    if cnpj0 == cnpj3:
        print("Todas Notas Fiscais Web Glamour")
        todas_notas_webglamour = pd.concat((lista_nf_df0, lista_nf_df3), ignore_index=True)
        todas_notas_webglamour.to_excel("Todas_Notas_Web_Glamour.xlsx")
        print(todas_notas_webglamour)
        
        
if cnpj1 == ['']:
    if cnpj1 == cnpj2:
        print("Todas Notas Fiscais Mamoeiro")
        todas_notas_Mamoeiro = pd.concat((lista_nf_df1, lista_nf_df2), ignore_index=True)
        todas_notas_Mamoeiro.to_excel("Todas_Notas_Mamoeiro.xlsx")
        print(todas_notas_webglamour)
        
        
if cnpj1 ==  ['']:
    if cnpj1 == cnpj3:
        print("Todas Notas Fiscais Mamoeiro")
        todas_notas_Mamoeiro = pd.concat((lista_nf_df1, lista_nf_df3), ignore_index=True)
        todas_notas_Mamoeiro.to_excel("Todas_Notas_Mamoeiro.xlsx")
        print(todas_notas_webglamour)
