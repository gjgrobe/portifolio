from tkinter import *
import pandas as pd
from tkinter import filedialog as dlg
from sqlalchemy.engine import URL
from sqlalchemy import create_engine
import pyodbc

def enginebd():
    conn_str = (
        r'DRIVER={ODBC Driver 11 for SQL Server};'
        r'SERVER=localhost\supera;'
        r'DATABASE=SGE;'
        r'UID=sa;'
        r'PWD=Fime2404;'
        r'Trusted_Connection=yes;'
    )

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conn_str})
    engine = create_engine(connection_url)
    return engine

def lerdados():
    caminho = dlg.askopenfilename(filetypes=[("Excel file","*.xlsx"),("Excel file 97-2003","*.xls")])
    return (caminho)

def atualiza():
    caminho = lerdados()
    df = pd.read_excel(caminho)
    df.dropna(inplace = True) #limpa os dados
    df.select_dtypes(include=['int', 'float']) #seleciona int e float
    rows, collumns = df.shape
    if collumns == 3:
        engine = enginebd()
        resposta['text'] = "AGUARDE OS DADOS DO PREVIROSA, ATUALIZANDO ..."
        j = 0
        for i in range (len(df)):
            matricula = df['Funcionário/Contrato'].str.split("-", n = 0, expand = True).iloc[i]
            matricula = matricula[0]
            valor = df['Total por Funcionário'].iloc[i]
            query = "UPDATE FATURA_SERVICOS_MENSAIS SET VALOR = ? WHERE IDRelacao IN (SELECT IDEmpresa FROM EMPRESAS WHERE OrgaoPF = ?) AND IDProdServ IN (SELECT IDServico FROM SERVICOS_PRODUTOS WHERE FixoR = 1)"
            registro = engine.execute(query, valor, matricula).rowcount
            j = registro + j
        resposta['text'] = ("Foram atualizados "+ (str(j)) +" registros de um total de " + (str(len(df))) + " ref. ao previrosa.")
    elif collumns == 4:
        engine = enginebd()
        resposta['text'] = "AGUARDE OS DADOS DA PREFEITURA, ATUALIZANDO ..."
        j = 0
        for i in range (len(df)):
            matricula = df['MUNICIPIO DE SANTA ROSA'].str.split("-", n = 0, expand = True).iloc[i]
            matricula = matricula[0]
            valor = df['Unnamed: 3'].iloc[i]
            query = "UPDATE FATURA_SERVICOS_MENSAIS SET VALOR = ? WHERE IDRelacao IN (SELECT IDEmpresa FROM EMPRESAS WHERE OrgaoPF = ?) AND IDProdServ IN (SELECT IDServico FROM SERVICOS_PRODUTOS WHERE FixoR = 1)"
            registro = engine.execute(query, valor, matricula).rowcount
            j = registro + j
        resposta['text'] = ("Foram atualizados "+ (str(j)) +" registros de um total de " + (str(len(df))) + " ref. a prefeitura.")
    else :
        resposta['text'] = ('O arquivo não corresponde ao layout correto, selecione outro arquivo')
    pass

def atualizacc():
    engine = enginebd()
    query1 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 4 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 5 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 44)"
    query2 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 5 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 4 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 44)"
    query3 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 6 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO IN (1,2,3) AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 44)"
	
    query4 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 21 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 5 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 43)"
    query5 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 22 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 4 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 43)"
    query6 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 23 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO IN (1,2,3) AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 43)"
	
    query7 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 13 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 5 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 42)"
    query8 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 12 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 4 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 42)"
    query9 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 14 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO IN (1,2,3) AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 42)"
	
    query10 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 30 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 5 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 45)"
    query11 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 31 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO = 4 AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 45)"
    query12 = "UPDATE FATURA_ITENS_ABERTO SET IDCC = 32 WHERE IDProdServ IN (SELECT IDSERVICO FROM SERVICOS_PRODUTOS) AND IDPLANO IN (1,2,3) AND IDRelacao IN (SELECT IDEMPRESA FROM EMPRESAS WHERE IDTipoAssociacao = 45)"
    registro1 = engine.execute(query1).rowcount
    registro2 = engine.execute(query2).rowcount
    registro3 = engine.execute(query3).rowcount
    registro4 = engine.execute(query4).rowcount
    registro5 = engine.execute(query5).rowcount
    registro6 = engine.execute(query6).rowcount
    registro7 = engine.execute(query7).rowcount
    registro8 = engine.execute(query8).rowcount
    registro9 = engine.execute(query9).rowcount
    registro10 = engine.execute(query10).rowcount
    registro11 = engine.execute(query11).rowcount
    registro12 = engine.execute(query12).rowcount
    total = registro1 + registro2 + registro3 + registro4 + registro5 + registro6 + registro7 + registro8 + registro9 + registro10 + registro11 + registro12
    resposta['text'] = ('UM TOTAL DE '+ str(total) + ' REGISTROS FORAM ATUALIZADOS COM SUCESSO')
    pass





janela = Tk()
janela.title("ATUALIZA VALORES DAS MENSALIDADES")
janela.geometry("1024x600")

# Labels

informacao = Label(janela, font="Arial 12", justify=CENTER,
                   text="FAÇA UM BACKUP ANTES!")

informacao2 = Label(janela, font="Arial 12", justify=CENTER,
                    text="APÓS CARREGAR O ARQUIVO, OS VALORES SERÃO ATUALIZADOS")
resposta = Label(janela, font="Arial 14", justify=CENTER, text="...")

# botao de ação
btatualiza = Button(janela, text='1 - ATUALIZA MENSALIDADES', command=lambda: atualiza())
btajustacc = Button(janela, text='2 - ATUALIZA CENTRO CUSTOS', command=lambda: atualizacc())

# Posições
informacao.place(width=700, height=50, x=150, y=5)
informacao2.place(width=700, height=50, x=200, y=60)
btatualiza.place(width=400, height=50, x=300, y=120)
btajustacc.place(width=400, height=50, x=300, y=200)
resposta.place(width=700, height=50, x=150, y=400)


janela.mainloop()
