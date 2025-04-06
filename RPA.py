# -- coding: utf-8�--
import pandas as pd
import pyautogui
import time
from openpyxl import Workbook
import os
from pathlib import Path
import keyboard


def ler_tarefas(caminhotarefa):
    tarefa = pd.read_csv(caminhotarefa)
    return tarefa


def validar_nome(nome_arquivo):
    if len(nome_arquivo) < 1:
        return "padrao"
    caracteres_invalidos = "/ * ? < > | :".split()
    for i in caracteres_invalidos:
        if i in nome_arquivo:
            return False
    else:
        return nome_arquivo + '.txt'


def executar_tarefa(valor, nomeTarefa, nome_arquivo):
    time.sleep(5)
    if nomeTarefa == 'Escrever':
        try:
            with open(nome_arquivo, "w") as arquivo:
                arquivo.write(valor)
            return True
        except:
            return False


def abrir_arquivo(nome_arquivo):
    caminho_arquivo = os.path.abspath(nome_arquivo)
    pyautogui.hotkey('win', 'r')
    print(caminho_arquivo)
    time.sleep(2)
    keyboard.write(caminho_arquivo)

    time.sleep(2)
    pyautogui.press('enter')


def automacao(caminhoTarefa, caminhoRelatorio):
    tarefas = ler_tarefas(caminhoTarefa)
    resultados = []
    nome_arquivo, texto, nomeTarefa = '', '', ''
    for _, tarefa in tarefas.iterrows():
        nomeTarefa = tarefa['Tarefa']
        dadoTarefa = tarefa['Dado']

        print(f"Executando tarefa: {nomeTarefa}")
        time.sleep(2)

        try:
            if nomeTarefa == 'ValidarNome':
                nome_arquivo = validar_nome(dadoTarefa)

            if nomeTarefa == 'Escrever':
                texto = dadoTarefa

                valid = executar_tarefa(texto, nomeTarefa, nome_arquivo)
                result = "Sucesso" if valid else "Falha"
                resultados.append([nomeTarefa, result])
                continue

            if nomeTarefa == "Abrir":
                abrir_arquivo(nome_arquivo)
                result = "Sucesso" if valid else "Falha"
                resultados.append([nomeTarefa, result])
                break

            result = "Sucesso" if nome_arquivo else "Falha"
            resultados.append([nomeTarefa, result])
        except:
            resultados.append([nomeTarefa, 'Falha'])

    gerar_relatorio(caminhoRelatorio, resultados)


def gerar_relatorio(caminho_relatorio, resultados):
    try:
        wb = Workbook()
        ws = wb.active
        ws.append(['Tarefa', 'Resultado'])

        for resultado in resultados:
            ws.append(resultado)

        wb.save(caminho_relatorio)
        print(f"Relatório salvo em {caminho_relatorio}")
    except Exception as e:
        print(f"Erro ao gerar o relatório: {e}")


automacao('tarefas.csv', 'relatorio_tarefas.xlsx')