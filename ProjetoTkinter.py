from __future__ import annotations
import json
import locale
import tkinter as tk
from tkinter import ANCHOR, ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import requests
from datetime import datetime, timedelta
import numpy as np
from tqdm import tqdm
from tkinter import *
from tkinter.ttk import *

# aqui faremos uma req na API para pegar a lista de todas as moedas e carregar no combobox
requisicao = requests.get('https://economia.awesomeapi.com.br/json/all/')
dicionario_moedas = requisicao.json()
lista_moedas = list(dicionario_moedas.keys())

janela = tk.Tk()
janela.title('Ferramenta de Cotações de Moedas')

# aqui vamos criar a variavel que irá controlar o percentual da barra de progresso
var_barra = DoubleVar()

# função que irá pegar a cotação única de moeda
def pegar_cotacao():
    moeda = combobox_selecionarmoeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_textocotacao['text'] = f"A cotação da {moeda} no dia {data_cotacao} foi de R$ {valor_moeda}"

# função que irá pedir ao usuario para selecionar a planilha com as moedas a serem cotadas
def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:  
        label_arquivoselecionado['fg'] = 'green'
        label_arquivoselecionado['text'] = f'Arquivo Selecionado: {caminho_arquivo}'
 
# função que pegas as cotações das moedas através da primeira coluna da planilha
# Obs (tive que fazer um ajuste, pois o link de intervalo de datas do Awesome API não estava funcionando)
def atualizar_cotacoes():
    try:
        # aqui faremos a barra de progresso aparecer
        barra_progresso = ttk.Progressbar(janela, variable=var_barra, orient=HORIZONTAL, maximum=100)
        barra_progresso.grid(row=11, column=0, sticky='nsew', columnspan=3, padx=10, pady=10)

        # ler o df de moedas
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:,0]
        # pegas a data de de inicio e data de fim das cotacoes
        data_inicial = calendario_datainicial.get()
        data_final = calendario_datafinal.get()
        
        # vamos calcular a diferença entre as datas
        d2 = datetime.strptime(data_final, '%d/%m/%Y')
        d1 = datetime.strptime(data_inicial, '%d/%m/%Y')

        #diferença entre data inicial e final (somamos 1 pois no laço ele começa a contar a partir de zero)
        qnt_dias = abs((d2 - d1).days) + 1

        # aqui vamos calcular o percentual que irá ser incrementado na variavel da barra de progresso
        var_percent = len(moedas)
        calc_perc = 100 / var_percent
        soma_percent = 0
        
        # para cada moeda
        for moeda in moedas:
            # vamos limpar a mensagem de sucesso caso o usuario utilize a funcao mais de uma vez
            label_atualizarcotacoes['text'] = ""

            lista_cotacoes = []

            # aqui incrementamos o valor da variavel da barra de progresso
            soma_percent += calc_perc
            var_barra.set(soma_percent)

            # aqui criei um laço para fazer uma requisição por vez com o link de cotação de moeda única (*ajuste)
            for i in range(qnt_dias):

                # vamos pegar a primeira data e somar 1 dia a cada volta do laço
                soma_dia = d1 + timedelta(days=i)
                soma_dia = datetime.strftime(soma_dia, '%d/%m/%Y')

                # destrinchando a data para usar no link
                ano = soma_dia[-4:]
                mes = soma_dia[3:5]
                dia = soma_dia[:2]

                link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?' \
                            f'start_date={ano}{mes}{dia}&' \
                            f'end_date={ano}{mes}{dia}'
                
                requisicao_moeda = requests.get(link)
                cotacoes = requisicao_moeda.json()

                # vamos iterar sobre a lista que foi retornada do link, e pegar o bid e o timestamp
                for cotado in cotacoes:
                    timestamp = int(cotado['timestamp'])
                    bid = float(cotado['bid'])
                    # aqui ele irá adicionar todas as cotações (por datas) em uma lista
                    lista_cotacoes.append({'bid': bid, 'timestamp': timestamp})
                    # a cada loop renderiza a janela para não travar o programa
                    janela.update()
        
            # agora vamos iterar a lista que criamos no 'for' anterior pra criar a planilha com o pandas
            for cotacao in lista_cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                # criar uma coluna em um novo df com todas as cotacoes daquela moeda
                if data not in df:
                    df[data] = np.nan
                # preenchemos o bid de cada moeda(linha) na coluna de cada data respectiva
                df.loc[df.iloc[:, 0] == moeda, data] = f'R$ {bid:.2f}'
        # criar um arquivo com todas as cotacoes
        df.to_excel('Cotações_Moedas.xlsx')
        label_atualizarcotacoes['fg'] = 'green'
        label_atualizarcotacoes['text'] = "Arquivo Atualizado com Sucesso"
        # e atualizamos a janela
        janela.update()
        # remove a barra de progresso apos completar o ciclo
        barra_progresso.grid_remove()
    except:
        # caso o usuario tenha pego um arquivo de formato incorreto
        label_atualizarcotacoes['fg'] = 'red'
        label_atualizarcotacoes['text'] = "Selecione um arquivo Excel no Formato Correto"


# Cotação de 1 Moeda

label_cotacaomoeda = tk.Label(text="Cotação de 1 moeda específica", borderwidth=2, relief='solid', fg='white', bg='black')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionarmoeda = tk.Label(text="Selecionar Moeda", anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

label_selecionardia = tk.Label(text="Selecione o dia que deseja pegar a cotação", anchor='e')
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendario_moeda = DateEntry(year=2022, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nsew')

label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_pegarcotacao = tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

# Cotação de Várias Moedas

label_cotacaovariasmoedas = tk.Label(text="Cotação de Múltiplas Moedas", borderwidth=2, relief='solid', bg='black', fg='white')
label_cotacaovariasmoedas.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionararquivo = tk.Label(text="Selecione um arquivo em Excel com as Moedas na coluna A", anchor='e')
label_selecionararquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text="Clique para selecionar", command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arquivoselecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e', fg='red')
label_arquivoselecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_datainicial = tk.Label(text="Data Inicial", anchor='e')
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')

label_datafinal = tk.Label(text="Data Final",  anchor='e')
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')

calendario_datainicial = DateEntry(year=2022, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky='nsew' )

calendario_datafinal = DateEntry(year=2022, locale='pt_br')
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nsew' )

botao_atualizarcotacoes = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=1, padx=10, pady=10, sticky='nsew')

label_atualizarcotacoes = tk.Label(text="")
label_atualizarcotacoes.grid(row=10, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()