import pandas as pd
import openpyxl
import os
import win32com.client as win32
import time
import tkinter as tk
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from email.message import EmailMessage
import smtplib
import AutoUpdate

AutoUpdate.set_url("https://github.com/AlyssonKrombauer/timlista/blob/c39f531a19d195e931887cdb3d93e16d009ef022/Fatura%20celular%20v2.py")
AutoUpdate.set_current_version("0.0.2")

lista = pd.read_csv(r"lista.txt", sep=";")
filiais = pd.read_excel(r"filiais.xlsx")
classif = pd.read_excel(r"filiais.xlsx", sheet_name='class')
#retira os caracteres da coluna
lista['NumAcs'] = lista['NumAcs'].str.replace('-','', regex=False)

#converte as colunas que de object para int
lista["NumAcs"] = pd.to_numeric(lista["NumAcs"], errors="coerce")

#procv com as duas planilhas
lista = pd.merge(lista, filiais, on=['NumAcs'], how='right')

#organiza as colunas
lista = lista[['MesRef', 'Filial', 'Colaborador', 'NumAcs', 'Nome', 'Tpserv', 'Data', 'Hora', 'Origem', 'Destino',
               'NumCham', 'Duração', 'Valor', 'Email']]
#pasta atual
atual = os.getcwd()

#fatura jurandir e alvaro

#jurandir = lista.loc[lista['Colaborador'].str.contains('JURANDIR')]
alvaro = lista.loc[lista['Colaborador'].str.contains('ALVARO')]

#jurandir.to_excel('Jurandir.xlsx',sheet_name='Fatura', index=False)
alvaro.to_excel('Alvaro.xlsx',sheet_name='Fatura', index=False)

#criando o classificação
soma_filial = lista.groupby('Filial')['Valor'].sum()

#concatenar
classifa = classif.merge(soma_filial, left_on='CC', right_on='Filial', how='right')
classifa = classifa[['CC', 'Filial', 'Classificação', 'Valor']]
#soma por filial
valor_y_total = classifa['Valor'].sum()
#linha total
nova_linha = {'CC': '','Filial': '','Classificação': 'Total','Valor' : valor_y_total}
classifa = pd.concat([classifa, pd.DataFrame(nova_linha, index=[len(classifa)])])
classifa.to_excel('Classifica.xlsx',sheet_name='rateio', index=False)

#criando uma instância do objeto Workbook e um objeto Worksheet associado
workbook = Workbook()
worksheet = workbook.active

#escrevendo o DataFrame no Worksheet
for r in dataframe_to_rows(classifa, index=False, header=True):
    worksheet.append(r)

#definindo a largura das colunas A e D
    worksheet.column_dimensions['B'].width = 35
    worksheet.column_dimensions['C'].width = 12
    worksheet.column_dimensions['D'].width = 12
    
#salvando o arquivo
    workbook.save('Classifica.xlsx')

def opcao1():
#grupar os valores por colaborador e somar a coluna "valor"
    soma_valores = lista.groupby('Colaborador')['Valor'].sum()

#filtrar os colaboradores soma dos valores é maior que 3
    colaboradores_maior_0 = soma_valores[soma_valores > 0].index
#criar a pasta "todos" caso ela ainda não exista
    if not os.path.exists('todos'):
        os.makedirs('todos')
        
#iterar sobre os colaboradores cuja soma dos valores é maior que 3 e criar um arquivo xlsx
    for colaborador in colaboradores_maior_0:
        df = lista[lista['Colaborador'] == colaborador]
        nome_arquivo = f'{colaborador}.xlsx'
        caminho_arquivo = os.path.join('todos', nome_arquivo)
        df.to_excel(caminho_arquivo, index=False)        
def opcao2():
#filtro para chamadas
    chamadas = lista.loc[lista['Tpserv'].str.startswith('Chamadas')]

#agrupamento por colaborador e soma dos valores de chamadas
    soma_chamadas = chamadas.groupby('Colaborador')['Valor'].sum()

    # Filtro para colaboradores com soma de chamadas maior que 3
    colaboradores_maior_3 = soma_chamadas[soma_chamadas > 3].index

#criação da pasta "valor maior" se ela ainda não existir
    if not os.path.exists('valor maior'):
        os.makedirs('valor maior')
#iteração sobre os colaboradores cuja soma das chamadas é maior que 3
    for colaborador in colaboradores_maior_3:
#filtro para as chamadas do colaborador
        chamadas_colab = chamadas.loc[chamadas['Colaborador'] == colaborador]
#criação do arquivo xlsx com as chamadas do colaborador
        nome_arquivo = f'{colaborador}.xlsx'
        caminho_arquivo = os.path.join('valor maior', nome_arquivo)
        chamadas_colab.to_excel(caminho_arquivo, index=False)
mes = ""
entrega = ""
login = ""
senha = ""
def opcao3():
    def obter_mes():
        global mes
        global entrega
        global login_value
        global senha_value

        login_value = entry_login.get()
        senha_value = entry_senha.get()

        mes = entry_mes.get()
        entrega = entry_entrega.get()
        janela_mes.destroy()
        mandar_email()

    janela_mes = tk.Tk()
    janela_mes.geometry("300x200")

    label_janela_login = tk.Label(janela_mes, text="Seu e-mail:")
    label_janela_login.pack()
    entry_login = tk.Entry(janela_mes, width=30)
    entry_login.pack()

    label_janela_senha = tk.Label(janela_mes, text="Senha do e-mail")
    label_janela_senha.pack()
    entry_senha = tk.Entry(janela_mes, width=30)
    entry_senha.pack()

    label_janela_mes = tk.Label(janela_mes, text="Fatura referente ao mês de:")
    label_janela_mes.pack()
    entry_mes = tk.Entry(janela_mes, width=30)
    entry_mes.pack()

    label_janela_entrega = tk.Label(janela_mes, text="Responder até dia:")
    label_janela_entrega.pack()
    entry_entrega = tk.Entry(janela_mes, width=30)
    entry_entrega.pack()

    botao_enviar = tk.Button(janela_mes, text="Enviar", command=obter_mes)
    botao_enviar.pack()
    janela_mes.mainloop()

def mandar_email():
    chamadas = lista.loc[lista['Tpserv'].str.startswith('Chamadas')]
#agrupamento por colaborador e soma dos valores de chamadas
    soma_chamadas = chamadas.groupby('Colaborador')['Valor'].sum()
#filtro para colaboradores com soma de chamadas maior que 3
    colaboradores_maior_3 = soma_chamadas[soma_chamadas > 3].index
#criação da pasta "valor maior" se ela ainda não existir
    if not os.path.exists('valor maior'):
        os.makedirs('valor maior')
#iteração sobre os colaboradores cuja soma das chamadas é maior que 3
    for colaborador in colaboradores_maior_3:        
#filtro para as chamadas do colaborador
        chamadas_colab = chamadas.loc[chamadas['Colaborador'] == colaborador]    
#criação do arquivo xlsx com as chamadas do colaborador
        nome_arquivo = f'{colaborador}.xlsx'
        caminho_arquivo = os.path.join('valor maior', nome_arquivo)
        chamadas_colab.to_excel(caminho_arquivo, index=False)        
#envio do e-mail para o colaborador
        email_colab = lista.loc[lista['Colaborador'] == colaborador, 'Email'].iloc[0]

        sender = login_value
        recipient = email_colab
        message = f"""Prezado(a) {colaborador},

Segue em anexo planilha atualizada da conta telefônica, referente ao mês de {mes}.
Gentileza responder esse e-mail até dia {entrega} informando a existência ou não de ligações particulares, e seu valor total se houver no corpo do email, para que possamos providenciar o desconto e encaminhar o relatório à Diretoria.

Lembre-se de utilizar o código de operadora da TIM (41) para realizar ligações de longa distância.
"""     
        email = EmailMessage()
        email["From"] = sender
        email["To"] = recipient
        email["Subject"] = "CONTA DE TELEFONE CELULAR"
        email.set_content(message)
        attachment = os.path.join(atual + '\\valor maior\\' + nome_arquivo)

        with open(attachment, "rb") as f:
            email.add_attachment(f.read(),
                                 filename= nome_arquivo,
                                 maintype="application",
                                 subtype="xlsx"
                                 )
            
        smtp = smtplib.SMTP("smtp.office365.com", port=587)
        smtp.starttls()
        smtp.login(sender, senha_value)
        smtp.sendmail(sender, recipient, email.as_string())
        smtp.quit()
        
        time.sleep(1)        
#filtro para chamadas
janela = tk.Tk()
janela.geometry("280x150")
janela.title("Fatura celular")
label_janela = tk.Label(janela, text="Escolha uma opção\n")
label_janela.pack()

opcao1_btn = tk.Button(janela, text="Gerar todas as faturas", command=opcao1)
opcao1_btn.pack()

opcao2_btn = tk.Button(janela, text="Gerar os maiores que R$3,00", command=opcao2)
opcao2_btn.pack()

opcao3_btn = tk.Button(janela, text="Enviar e-mail dos maiores", command=opcao3)
opcao3_btn.pack()

janela.mainloop()
