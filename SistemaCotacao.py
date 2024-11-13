import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
from datetime import datetime
import requests

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = [moeda for moeda in dicionario_moedas]

def pegar_cotacao():
    moeda_selecionada = combox_selecionar_moeda.get()
    data_selecionada = calendario_moeda.get()
    ano = data_selecionada[-4:]
    mes = data_selecionada[3:5]
    dia = data_selecionada[:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda_selecionada}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao = requests.get(link)
    requisicao = requisicao.json()
    cotacao = requisicao[0]['bid']
    label_textocotacao['text'] = f'{moeda_selecionada}: {cotacao}'
    
def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o arquivo de moeda')
    var_caminho.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f'Arquivo selecionado: {caminho_arquivo}'

def atualizar_cotacoes():
    df = pd.read_excel(var_caminho.get())
    moedas = df.iloc[:, 0]
    data_inicial = calendario_datainicial.get()
    data_final = calendario_datafinal.get()

    ano_inicial = data_inicial[-4:]
    mes_inicial = data_inicial[3:5]
    dia_inicial = data_inicial[:2]
    
    ano_final = data_final[-4:]
    mes_final = data_final[3:5]
    dia_final = data_final[:2]

    for moeda in moedas:
        link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}'
        requisicao = requests.get(link)
        cotacoes = requisicao.json()
        for cotacao in cotacoes:
            timestamp = int(cotacao['timestamp'])
            bid = float(cotacao['bid'])
            data = datetime.fromtimestamp(timestamp)
            data = data.strftime('%d/%m/%Y')
            if data not in df:
                df[data] = np.nan

            df.loc[df.iloc[:, 0] == moeda, data] = bid
            label_atualizacao['text'] = 'Arquivo atualizado com sucesso'


janela = tk.Tk()

janela.title('Ferramenta de Cotação de Moedas')

label_cotacaomoeda = tk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
label_cotacaomoeda.grid(row=0, column=0, sticky='nsew', padx=10, pady=10, columnspan=3)

label_selecionarmoeda = tk.Label(text='Selecionar moeda')
label_selecionarmoeda.grid(row=1, column=0, sticky='nsew', padx=10, pady=10, columnspan=2)

combox_selecionar_moeda = ttk.Combobox(values=lista_moedas)
combox_selecionar_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

label_selecionardata = tk.Label(text='Selecionar data desejada')
label_selecionardata.grid(row=2, column=0, sticky='nsew', columnspan=2, padx=10, pady=10)

calendario_moeda = DateEntry(year=2024, locale='pt_br')
calendario_moeda.grid(column=2, row=2, padx=10, pady=10, sticky='nsew')

label_textocotacao = tk.Label(text='')
label_textocotacao.grid(row=3, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

botao_cotacao = tk.Button(text='Pegar Cotação', command=pegar_cotacao)
botao_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

# Várias Moedas
label_cotacaovariasmoedas = tk.Label(text='Cotação de várias moedas', borderwidth=2, relief='solid')
label_cotacaovariasmoedas.grid(row=4, column=0, sticky='nsew', padx=10, pady=10, columnspan=3)

label_selecionararquivo = tk.Label(text='Selecione um arquivo em Excel com as moedas na coluna A:')
label_selecionararquivo.grid(row=5, column=0, padx=10, pady=10, sticky='nsew', columnspan=2)

botao_selecionararquivo = tk.Button(text='Clique para selecionar arquivo', command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

var_caminho = tk.StringVar()

label_arquivoselecionado = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
label_arquivoselecionado.grid(row=6, column=0, padx=10, pady=10, sticky='nsew', columnspan=3)

label_datainicial = tk.Label(text='Data Inicial:')
label_datafinal = tk.Label(text='Data Final:')
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')

calendario_datainicial = DateEntry(year=2024, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, sticky='nsew', padx=10, pady=10)

calendario_datafinal = DateEntry(year=2024, locale='pt_br')
calendario_datafinal.grid(row=8, column=1, sticky='nsew', padx=10, pady=10)

botao_atualizarcotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizacao = tk.Label(text='')
label_atualizacao.grid(row=9, column=1, padx=10, pady=10, sticky='nsew', columnspan=2)

botao_fechar = tk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()