from datetime import date
import pandas as pd
from xlsxwriter import Workbook
from sqlalchemy.engine import URL
from sqlalchemy import Column, PrimaryKeyConstraint, create_engine
from sqlalchemy import insert, select, update
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy import Integer, String, DateTime, Float
import pyodbc
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkcalendar import Calendar, DateEntry
from datetime import datetime
from tkinter import filedialog as dlg


# pyinstaller --onefile --noconsole --hidden-import babel.numbers .\atualizaservicosmensais\descontosnaoduplicado.py

def enginebd():
    conn_str = (
        r'DRIVER={ODBC Driver 11 for SQL Server};'
        r'SERVER=localhost\supera;'
        r'DATABASE=SGE;'
        r'UID=******;' #OCULTADO O LOGIN, TROCAR QUANDO FOR UTILIZAR.
        r'PWD=******;' #OCULTADO A SENHA, TROCAR QUANDO FOR UTILIZAR.
        r'Trusted_Connection=yes;'
    )

    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": conn_str})
    engine = create_engine(connection_url)
    return engine

Base = declarative_base()

class NOTAS_PARCELAS_RECEITAS_TIPO_RECEBIMENTO(Base):
    '''
    Classe para a declaração da tabela
    
    ID = Column(Integer(), primary_key = True)
    
    IDVencimento = Column(Integer())
    
    TipoRecebimento = Column(Integer())
    
    ChequeAgencia = Column(String(15))
    
    ChequeConta	 = Column(String(20))
    
    ChequeNumero = Column(String(30))
    
    ChequePreDatado = Column(Integer())
    
    ChequeBomPara = Column(DateTime())
    
    Compensado = Column(Integer())
    
    CompensadoEm = Column(DateTime())
    
    CodiElo	= Column(Integer())
    
    Valor = Column(Float())
    
    IDBanco	= Column(Integer())
    
    IDCartao = Column(Integer())
    
    Historico = Column(String(100))
    
    ChequeUtilizadoPara = Column(String(100))
    
    ChequeEmitente = Column(String(70))
    
    ChequeTroco = Column(Float())
    
    DataMovimento = Column(DateTime())
    
    ChequeTipo = Column(Integer())
    
    ChequeExportarTroca = Column(Integer())
    
    Observacoes = Column(String(200))
    
    IDCHEQUE_ORIGEM = Column(Integer())
    '''
    __tablename__ = 'NOTAS_PARCELAS_RECEITAS_TIPO_RECEBIMENTO'
    ID = Column(Integer, primary_key = True, autoincrement = False)
    IDVencimento = Column(Integer())
    TipoRecebimento = Column(Integer())
    ChequeAgencia = Column(String(15))
    ChequeConta	 = Column(String(20))
    ChequeNumero = Column(String(30))
    ChequePreDatado = Column(Integer())
    ChequeBomPara = Column(DateTime())
    Compensado = Column(Integer())
    CompensadoEm = Column(DateTime())
    CodiElo	= Column(Integer())
    Valor = Column(Float())
    IDBanco	= Column(Integer())
    IDCartao = Column(Integer())
    Historico = Column(String(100))
    ChequeUtilizadoPara = Column(String(100))
    ChequeEmitente = Column(String(70))
    ChequeTroco = Column(Float())
    DataMovimento = Column(DateTime())
    ChequeTipo = Column(Integer())
    ChequeExportarTroca = Column(Integer())
    Observacoes = Column(String(200))
    IDCHEQUE_ORIGEM = Column(Integer())

class descontosAssociados:
    def __init__(self) -> None:
        pass

    def carregaArquivo(self):
        link = dlg.askopenfilename(filetypes=[("Excel 97-2003",r"*.xlsx .xls")])
        link = str(link)
        return (link)

    def lerDesconto(self):
        dados = self.carregaArquivo()
        #dados = ("./atualizaservicosmensais/Descontos Sindicato.xls")
        df = pd.read_excel(dados)
        if 'MUNICIPIO DE SANTA ROSA' in df.columns:
            dfnew = df[['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 26']].copy()
            dfnew.rename(columns = {'Unnamed: 0':'MATRICULA', 'Unnamed: 3':'NOME', 'Unnamed: 26':'VALOR'}, inplace = True)
            dfnew = dfnew.groupby(['NOME']).agg({'MATRICULA': lambda x: list(set(x)), 'NOME': lambda x: list(set(x)), 'VALOR':'sum'})
        elif 'INSTITUTO DE PREVIDENCIA DOS SERVIDORES MUNICIPAIS' in df.columns:
            dfnew = df[['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 20']].copy()
            dfnew.rename(columns = {'Unnamed: 0':'MATRICULA', 'Unnamed: 3':'NOME', 'Unnamed: 20':'VALOR'}, inplace = True)
            dfnew = dfnew.groupby(['NOME']).agg({'MATRICULA': lambda x: list(set(x)), 'NOME': lambda x: list(set(x)), 'VALOR':'sum'})
        return (dfnew)

    def gravaDesconto(self, idFatura, dtQuitacao, IDBanco):
        df = self.lerDesconto()
        engine = enginebd()
        Session = sessionmaker(bind = engine)
        session = Session()
        qvalororiginal = "select ValorOriginal from NOTA_PARCELAS_RECEITAS WHERE IDNOTA IN ( SELECT IDNota FROM NOTAS_CABECALHO_RECEITAS WHERE idFatura = ? AND IDRelacao IN (SELECT IDEmpresa FROM EMPRESAS WHERE OrgaoPF = ?))"
        qidvencimento = "select idvencimento from NOTA_PARCELAS_RECEITAS WHERE IDNOTA IN ( SELECT IDNota FROM NOTAS_CABECALHO_RECEITAS WHERE idFatura = ? AND IDRelacao IN (SELECT IDEmpresa FROM EMPRESAS WHERE OrgaoPF = ?))"
        qdtQuitacao = "UPDATE NOTA_PARCELAS_RECEITAS SET DataRece = ? where idvencimento = ?"
        qdtQuitacaoAcrescimo = "UPDATE NOTA_PARCELAS_RECEITAS SET DataRece = ?, Acrescimo = ? where idvencimento = ?"
        qdtQuitacaoDesconto = "UPDATE NOTA_PARCELAS_RECEITAS SET DataRece = ?, Desconto = ? where idvencimento = ?"
        maxID = 'select max(id) from NOTAS_PARCELAS_RECEITAS_TIPO_RECEBIMENTO'
        qExisteIDVencimento = "select idvencimento from NOTAS_PARCELAS_RECEITAS_TIPO_RECEBIMENTO where idvencimento = ?"
        qAgencia = 'select NumeroAgencia from BANCOS where IDBanco = ?'
        qConta = 'select NumeroConta from BANCOS where IDBanco = ?'
        with engine.connect() as conn:
    
            result = (
            conn.
            execution_options(yield_per=100).
            engine.execute(qAgencia, IDBanco)
            )
            for partition in result.partitions():
                # ENCONTRA A AGENCIA PARA INCLUIR NA QUITACAO DA PARCELA
                for row in partition:
                    agencia = row[0]
            result = (
            conn.
            execution_options(yield_per=100).
            engine.execute(qConta, IDBanco)
            )
            for partition in result.partitions():
                # ENCONTRA A AGENCIA E CONTA PARA INCLUIR NA QUITACAO DA PARCELA
                for row in partition:
                    conta = row[0]
        listadiferentes = []
        valororiginal = -1
        contaNota = 0
        contaInconsistencias = 0
        contaDuplicado = 0
        existeIDVencimento = -1
        for i in range (len(df)):
            validaIDVencimento = False
            matricula = df[['MATRICULA']].iloc[i]
            matricula = str(matricula[0])
            matricula = matricula.split("-")
            matricula = matricula[0]
            matricula = matricula.split("""['""")
            matricula = matricula[1]
            nome = df[['NOME']].iloc[i]
            nome = str(nome[0])
            nome = nome.split("""['""")
            nome = nome[1]
            nome = nome.split("""']""")
            nome = nome[0]
            valor = float(df[['VALOR']].iloc[i])
            valor = round(valor, ndigits = 2)
            with engine.connect() as conn:
                
                result = (
                conn.
                execution_options(yield_per=100).
                engine.execute(qvalororiginal, idFatura, matricula)
                )
                for partition in result.partitions():
                    # ENCONTRA O VALOR ORIGINAL DA NOTA PARA COMPARAR COM O VALOR DO DESCONTO
                    for row in partition:
                        valororiginal = float((round(row[0], ndigits=2)))
            
            if valororiginal != -1 :
                result = (
                conn.
                execution_options(yield_per=100).
                engine.execute(qidvencimento, idFatura, matricula)
                )
                for partition in result.partitions():
                    # SELECIONA O IDVENCIMENTO PARA PODER SABER QUAL NOTA DEVE SER QUITADA
                    for row in partition:
                        idvencimento = row[0]

                result = (
                conn.
                execution_options(yield_per=100).
                engine.execute(qExisteIDVencimento, idvencimento)
                )
                for partition in result.partitions():
                    # IDENTIFICA SE JÁ EXISTE UMA IDVENCIMENTO NA TABELA NOTAS_PARCELAS_RECEITAS_TIPO_RECEBIMENTO
                    for row in partition:
                        existeIDVencimento = row[0]
                        #print ('Estou em duvida eh True?', existeIDVencimento)
                        if idvencimento == existeIDVencimento:
                            validaIDVencimento = True
                        else:
                            validaIDVencimento = False

                result = (
                conn.
                execution_options(yield_per=100).
                engine.execute(maxID)
                )
                for partition in result.partitions():
                    # SELECIONA O ULTIMO ID PARA INCREMENTAR
                    for row in partition:
                        ultimoID = row[0]

                if (valor == valororiginal) and (validaIDVencimento is False) :
                    id = ultimoID + 1
                    engine.execute(qdtQuitacao, dtQuitacao, idvencimento)
                    quitar = NOTAS_PARCELAS_RECEITAS_TIPO_RECEBIMENTO(ID = id, ChequeAgencia = agencia, ChequeConta = conta, ChequePreDatado = 0, IDVencimento = idvencimento, TipoRecebimento = 0,
                                                                    Compensado = 1, ChequeBomPara = dtQuitacao, CodiElo = 0, Valor = valor, CompensadoEm = dtQuitacao, IDBanco = IDBanco, IDCartao = 0,
                                                                    DataMovimento = dtQuitacao, ChequeTipo = 0, ChequeExportarTroca = 1, ChequeTroco = 0.00)
                    session.add(quitar)
                    session.commit()
                    contaNota = contaNota + 1
                elif (valor == valororiginal) and (validaIDVencimento is True):
                    contaDuplicado = contaDuplicado + 1
                    #print ('VALOR DUPLICADO NA BASE DE DADOS')
                else:
                    diferencav = (valor - valororiginal)
                    id = ultimoID + 1
                    if diferencav > 0 :
                        engine.execute(qdtQuitacaoAcrescimo, dtQuitacao, diferencav, idvencimento)
                    elif (diferencav < 0):
                        diferencav = (diferencav * -1)
                        engine.execute(qdtQuitacaoDesconto, dtQuitacao, diferencav, idvencimento)
                        diferencav = (diferencav * -1)
                    quitar = NOTAS_PARCELAS_RECEITAS_TIPO_RECEBIMENTO(ID = id, ChequeAgencia = agencia, ChequeConta = conta, ChequePreDatado = 0, IDVencimento = idvencimento, TipoRecebimento = 0,
                                                                    Compensado = 1, ChequeBomPara = dtQuitacao, CodiElo = 0, Valor = valor, CompensadoEm = dtQuitacao, IDBanco = IDBanco, IDCartao = 0,
                                                                    DataMovimento = dtQuitacao, ChequeTipo = 0, ChequeExportarTroca = 1, ChequeTroco = 0.00)
                    session.add(quitar)
                    session.commit()
                    inconsistencia = ("Inconsistência: ", i, matricula, nome , valor, valororiginal, diferencav)
                    listadiferentes.append(inconsistencia)
                    contaInconsistencias = contaInconsistencias + 1
            else:
                lbresposta['text'] = ('VALOR NÃO ENCONTRADO')
        valoresdiferentes = pd.DataFrame(listadiferentes)
        valoresdiferentes.rename(columns = {0:'Inconsistencia', 1:'Linha do Arquivo', 2:'Matricula', 3:'Nome', 4: 'VALOR DO DESCONTO', 5:'VALOR LANÇADO NO SUPERA', 6:'Diferenca (VALOR DESCONTO - VALOR SUPERA)'}, inplace = True)
        #print(valoresdiferentes.head(50))
        path = 'inconsistencia-' + str(datetime.now().strftime('%Y-%m-%d %H-%M-%S')) + '.xlsx'
        #valoresdiferentes.to_excel('inconsistencias.xlsx', engine='xlsxwriter', index = False)
        valoresdiferentes.to_excel(path, index = False)
        #valoresdiferentes.to_excel('./atualizaservicosmensais/inconsistencias.xlsx', index = False)
        lbresposta['text'] = 'Foram quitados com sucesso ' + str(contaNota) + ' registros e geradas ' + str(contaInconsistencias) + ' inconsistências e ' + str(contaDuplicado) + ' já quitados.'
        lbresposta2['text'] = 'Foi criado o arquivo ' + path + ' de inconsistências para conferência.'
    
    def cFatura(self):
        engine = enginebd()
        qFatura = "select DESCRICAO, IDFatura FROM FATURA_CABECALHO where idFatura > 0 order by FaturaNumero desc"
        cFatura = []
        with engine.connect() as conn:
                result = (
                conn.
                execution_options(yield_per=100).
                engine.execute(qFatura)
                )
                for partition in result.partitions():
                    # CRIA UMA LISTA COM AS FATURAS GERADAS PARA SELEÇÃO POSTERIOR
                    for row in partition:
                        cFatura.append(row)
        return (cFatura)

    def banco(self):
        engine = enginebd()
        qIDBanco = "select NumeroConta, idBanco from BANCOS where Inativo = 0 and IDBanco >= 0"
        listaBanco = []
        with engine.connect() as conn:
            result = (
            conn.
            execution_options(yield_per=100).
            engine.execute(qIDBanco)
            )
            for partition in result.partitions():
                # CRIA UMA LISTA COM AS CONTAS BANCARIAS PARA SELECIONAR
                for row in partition:
                    listaBanco.append(row)
        return (listaBanco)

root = tk.Tk()
root.title("LEITURA DOS DESCONTOS EM FOLHA")
root.geometry("1024x600")

dA = descontosAssociados()

#### CRIAÇÃO DA COMBO PARA SELECIONAR O FATURAMENTO
lista = dA.cFatura()

def get_index(*args):
     #print (varComboFatura.get()) #MOSTRA A DESCRIÇÃO DO ITEM SELECIONADO
     listaFatura = lista[comboFatura.current()][1] #PEGA A POSIÇÃO DA LISTA PARA ENCONTRAR O ID
     return (listaFatura)

varComboFatura = StringVar(value = 'SELECIONE UM FATURAMENTO:') #CRIA UMA VARIAVEL PARA ARMAZENAR O VALOR SELECIONADO
comboFatura = ttk.Combobox(root, textvariable=varComboFatura)
comboFatura['values'] = [lista[i][0] for i in range(len(lista))] #TRAS A LISTA PARA A COMBO E O FOR FAZ APARECER TODOS OS ITENS
comboFatura['state'] = 'readonly' #DEFINE A COMBO COMO SOMENTE LEITURA
comboFatura.place(width=700, height=50, x=150, y=20) #DEFINE A POSIÇÃO DA COMBO
varComboFatura.trace('w', get_index) #ENVIA PARA A FUNÇÃO GET_INDEX O VALOR

#### CRIAÇÃO DA COMBO PARA SELECIONAR O BANCO

listabanco = dA.banco()
#print (listabanco)

def get_index_banco(*args):
    idBanco = (listabanco[comboBanco.current()][1])
    return (idBanco)
varComboBanco = StringVar(value = 'SELECIONE UMA CONTA BANCÁRIA:')
comboBanco = ttk.Combobox(root, textvariable= varComboBanco)
comboBanco['values'] = [listabanco[i][0] for i in range(len(listabanco))]
comboBanco['state'] = 'readonly'
comboBanco.place(width=700, height=50, x=150, y=80)
varComboBanco.trace('w', get_index_banco)


#### CRIAÇÃO DO CALENDARIO PARA SELECIONAR A DATA DE RECEBIMENTO
def print_sel():
    data = cal.selection_get()
    data = str(data)
    return (data)

cal = Calendar(root, font="Arial 14", selectmode='day', locale='pt_BR', cursor="hand1")
cal.pack(pady = 150)

confirma = ttk.Button(root, text = 'RECEBER OS DESCONTOS', command = lambda: dA.gravaDesconto(idFatura = get_index(), dtQuitacao= print_sel(), IDBanco= get_index_banco()) )
confirma.place(width=700, height=50, x=150, y=410)

lbresposta = ttk.Label(root, font="Arial 12", text = '*IMPORTANTE! ANTES DE EXECUTAR A ROTINA, FAÇA UM BACKUP DO SISTEMA!*')
lbresposta.place(width=700, height=50, x=150, y=480)

lbresposta2 = ttk.Label(root, font="Arial 10", text = 'APÓS CARREGAR O ARQUIVO A ROTINA SERÁ EXECUTADA AUTOMATICAMENTE.')
lbresposta2.place(width=700, height=50, x=150, y=530)


root.mainloop()
