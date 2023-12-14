import PySimpleGUI as sg
import pandas as pd
import os
import random


def criar_janela():
    temas_disponiveis = [
        'Default', 'Dark', 'LightGrey', 'DarkGrey', 'BlueMono', 'BrownBlue', 'DarkAmber', 'DarkBlue', 'DarkBrown', 'DarkGreen',
        'DarkPurple', 'DarkTeal', 'Dark2', 'Dark5', 'Dark7', 'DarkBlack', 'Material1', 'Material2', 'Material3', 'Material4',
        'DarkRed1', 'DarkRed2', 'DarkRed3', 'DarkRed4', 'DarkRed5', 'DarkRed6', 'DarkRed7', 'DarkRed8', 'LightGreen', 'LightBlue',
        'LightPurple', 'LightOrange', 'LightPink', 'LightBrown', 'LightGray', 'LightYellow', 'LightCyan', 'LightTeal',
        'LightOliveGreen', 'LightGold', 'LightSilver', 'LightRed', 'LightMagenta', 'DarkGray', 'DarkYellow', 'DarkCyan',
        'DarkTeal2', 'DarkOliveGreen', 'DarkGold', 'DarkSilver'
    ]  # Lista de temas disponíveis
    tema_aleatorio = random.choice(temas_disponiveis)
    sg.theme(tema_aleatorio)

    layout = [
        [sg.Frame('', layout=[
            [sg.Text('Selecione pelo menos um arquivo .XLSX das Horas:',
                     font=('Sans Serif', 8), justification='center')]
        ], relief=sg.RELIEF_SUNKEN, pad=(0, 5))],

        [sg.Text('GMN:', size=(10, 1)), sg.Input(key='-GMN-', default_text='Arquivo de horas da GMN', enable_events=True),
         sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"),), button_text='Escolha o arquivo')],

        [sg.Text('UMAN:', size=(10, 1)), sg.Input(key='-UMAN-', default_text='Arquivo de horas da UMAN', enable_events=True),
         sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"),), button_text='Escolha o arquivo')],

        [sg.Text('UMEN:', size=(10, 1)), sg.Input(key='-UMEN-', default_text='Arquivo de horas da UMEN', enable_events=True),
         sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"),), button_text='Escolha o arquivo')],

        [sg.Text('UOAN:', size=(10, 1)), sg.Input(key='-UOAN-', default_text='Arquivo de horas da UOAN', enable_events=True),
         sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"),), button_text='Escolha o arquivo')],

        [sg.Text('UTEN:', size=(10, 1)), sg.Input(key='-UTEN-', default_text='Arquivo de horas da UTEN', enable_events=True),
         sg.FileBrowse(file_types=(("XLSX Files", "*.xlsx"),), button_text='Escolha o arquivo')],

        [sg.Text('Nome do arquivo de saída:', size=(20, 1)), sg.Input(
            key='-NOME-', default_text='Para_importacao_no_Protheus')],

        [sg.Column([[sg.Text('')]], pad=(0, 10))],

        [sg.Button('Processar', size=(10, 1), button_color=(
            'white', '#4CAF50'), font=('Helvetica', 12), border_width=2)],

        [sg.Column([[sg.Text('')]], pad=(0, 10))],

        [sg.Frame('', layout=[
            [sg.Text('Hagliberto Alves de Oliveira', font=(
                'Helvetica', 8), justification='center')]
        ], relief=sg.RELIEF_SUNKEN, pad=(0, 5))]
    ]
    # Cria a janela
    return sg.Window('Processar Arquivos XLSX', layout, element_justification='c')


def ler_e_processar_arquivo(arquivo):
    try:
        dados = pd.read_excel(arquivo, header=None)
        dados[0] = dados[0].apply(lambda x: f"{x:06}")
        dados[2] = dados[2].apply(lambda x: f"{x:.2f}".replace(",", "."))
        return dados
    except Exception as e:
        sg.popup_error(f'Erro na leitura do arquivo {arquivo}: {e}')
        return pd.DataFrame()


def processar_arquivos(arquivos, nome_arquivo_saida, pasta_destino):
    dados_completos = pd.DataFrame()

    for arquivo in arquivos:
        if arquivo:
            dados = ler_e_processar_arquivo(arquivo)
            dados_completos = pd.concat(
                [dados_completos, dados], ignore_index=True)

    if not dados_completos.empty:
        dados_completos = dados_completos[dados_completos[2] != "0.00"]

        if not dados_completos.empty:
            dados_completos.sort_values([0, 1], inplace=True)

            # Cria a pasta de destino se ela não existir
            if not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)

            # Salva o arquivo de saída na pasta destino
            arquivo_saida = os.path.join(
                pasta_destino, f'{nome_arquivo_saida}.xlsx')
            dados_completos.to_excel(arquivo_saida, index=False, header=False)

            arquivo_saida = f'{nome_arquivo_saida}.xlsx'
            dados_completos.to_excel(
                arquivo_saida, index=False, header=False)

            for verba in dados_completos[1].unique():
                dados_verba = dados_completos[dados_completos[1] == verba]
                nome_arquivo = os.path.join(
                    pasta_destino, f'Verba_{verba}_Tratada.xlsx')
                dados_verba.sort_values([0, 1], inplace=True)
                dados_verba.to_excel(nome_arquivo, index=False, header=False)
        else:
            sg.popup(
                'Erro!', 'Todos os dados possuem valor 0.00 na terceira coluna.')
    else:
        sg.popup('Erro!', 'Nenhum dado encontrado nos arquivos selecionados.')

    if not dados_completos.empty:
        dados_completos = dados_completos[dados_completos[2] != "0.00"]
        if not dados_completos.empty:
            dados_completos.sort_values([0, 1], inplace=True)

            # Cria a pasta de destino se ela não existir
            if not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)

            # Salva o arquivo de saída na pasta destino
            arquivo_saida = os.path.join(
                pasta_destino, f'{nome_arquivo_saida}.xlsx')
            dados_completos.to_excel(
                arquivo_saida, index=False, header=False)

            arquivo_saida = f'{nome_arquivo_saida}.xlsx'
            dados_completos.to_excel(
                arquivo_saida, index=False, header=False)

            for verba in dados_completos[0].unique():
                dados_verba = dados_completos[dados_completos[0] == verba]
                nome_arquivo = os.path.join(
                    pasta_destino, f'Horas_da_Matricula_{verba}.xlsx')
                dados_verba.sort_values([0, 1], inplace=True)
                dados_verba.to_excel(
                    nome_arquivo, index=False, header=False)
        else:
            sg.popup(
                'Erro!', 'Todos os dados possuem valor 0.00 na terceira coluna.')
    else:
        sg.popup('Erro!', 'Nenhum dado encontrado nos arquivos selecionados.')


janela = criar_janela()

while True:
    evento, valores = janela.read()

    if evento == sg.WINDOW_CLOSED:
        break
    elif evento == 'Processar':
        arquivos = [valores['-GMN-'], valores['-UMAN-'],
                    valores['-UMEN-'], valores['-UOAN-'], valores['-UTEN-']]
        arquivos = [arquivo for arquivo in arquivos if arquivo]
        nome_arquivo_saida = valores['-NOME-']
        pasta_destino = 'Arquivos Formatados'  # Substitua pelo caminho desejado

        if len(arquivos) > 0:
            processar_arquivos(arquivos, nome_arquivo_saida, pasta_destino)
            break
        else:
            sg.popup('Erro!', 'Por favor, selecione pelo menos um arquivo XLSX.')

janela.close()
