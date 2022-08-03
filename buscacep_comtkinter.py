from tkinter import *
from tkinter import ttk
import requests


def pega_cep(cep):
    '''
    FUNÇÃO QUE RETORNA ENDEREÇO, CIDADE E ESTADO, USANDO A CONSULTA PELO CEP.
    '''
    try:
        cep = cep.get()
        link = f"https://cep.awesomeapi.com.br/json/{cep}"
        requisicao = requests.get(link)
        dic_requisicao = requisicao.json()
        endereco = dic_requisicao["address"]
        cidade1 = dic_requisicao["city"]
        estado1 = dic_requisicao["state"]
        saida_rua ["text"] = endereco
        cidade ["text"] = cidade1
        estado["text"] = estado1

    except ValueError:

        saida_rua ["text"] = "ERRO"
        cidade ["text"] = "ERRO"
        estado["text"] = "ERRO"
        pass

janela = Tk()
janela.title ("Busca CEP")
janela.geometry("1024x600")

#Labels
informacao = Label(janela, font= "Arial 20", justify=CENTER, text = "Informe o CEP: ")
texto1 = Entry(janela, font= "Arial 20", justify=CENTER)
saida_rua = Label(janela,font = "Arial 20", text="A rua é: ")
cidade = Label(janela, font = "Arial 20", text = "A cidade é: ")
estado = Label(janela, font = "Arial 20", text = "O estado é: ")

#variaveis globais
cep = texto1

#Posições
informacao.place(width=700, height=50, x= 200, y= 5)
texto1.place(width=300, height=50, x= 400, y= 60)
saida_rua.place(width=800, height=50, x= 200, y= 230)
cidade.place(width=800, height=50, x=200, y = 270)
estado.place(width=300, height=50, x=400, y=320)

#botao de ação
botao = Button(janela, text="BUSCACEP", command = lambda:pega_cep(cep))

#posição do botao
botao.place(width=300, height=50, x= 400, y= 150)


janela.mainloop()
