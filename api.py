from tkinter import *
from tkinter import ttk
import requests
#from flask import Flask, jsonify


def btclick():
    enviainformacoesbd(nome, cpf)
    texto1.delete(first=0, last=1000)
    texto2.delete(first=0, last=1000)


def enviainformacoesbd(nome, cpf):
    '''
    FUNÇÃO RESPONSAVEL POR GRAVAR OS DADOS EM UM BANCO FIREBASE, ONLINE.
    '''
    nome = nome.get()
    cpf = cpf.get()
    variavel = '{' + '"Nome"' + ' : ' + '"' + nome + \
        '"' + ' , ' + '"CPF"' + ' : ' + '"' + cpf + '"}'
    requisicao = requests.post(
        "https://grobe-piloto3-default-rtdb.firebaseio.com/.json", data=variavel)
    if requisicao.status_code == 200:
        resposta['text'] = "Dados enviados com sucesso"
    else:
        resposta['text'] = "Problema na conexão"


def format_cpf(event=None):
    '''
    FUNÇÃO RESPONSAVEL POR FORMATAR O CEP 00000-000
    '''
    text = texto2.get().replace(".", "").replace("-", "")[:11]
    new_text = ""

    if event.keysym.lower() == "backspace":
        return

    for index in range(len(text)):
        if not text[index] in "0123456789":
            continue
        if index in [2, 5]:
            new_text += text[index] + "."
        elif index == 8:
            new_text += text[index] + "-"
        else:
            new_text += text[index]

    texto2.delete(0, "end")
    texto2.insert(0, new_text)


janela = Tk()
janela.title("ENVIA DADOS BD")
janela.geometry("1024x600")

# Labels
informacao = Label(janela, font="Arial 20", justify=CENTER,
                   text="INFORME O NOME E CPF: ")

texto1 = Entry(janela, font="Arial 20", text="NOME COMPLETO", justify=CENTER)

#FORMATA CAMPO COMO CPF
texto2 = Entry(janela, font="Arial 20", text="CPF", justify=CENTER)
texto2.bind("<KeyRelease>", format_cpf)
texto2.pack()

resposta = Label(janela, font="Arial 20", text="RESPOSTA DA REQUISICAO")


# variaveis que vão pra função
nome = texto1
cpf = texto2

# Posições
informacao.place(width=700, height=50, x=150, y=5)
texto1.place(width=400, height=50, x=300, y=60)
texto2.place(width=400, height=50, x=300, y=150)
resposta.place(width=700, height=50, x=150, y=230)

# botao de ação
botao = Button(janela, text="ENVIAR", command=lambda: btclick())

# posição do botao
botao.place(width=400, height=50, x=300, y=300)


janela.mainloop()
