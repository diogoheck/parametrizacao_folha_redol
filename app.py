import csv
from openpyxl import load_workbook
import os

class Rota:
    centro_custo = False
    evento = False
    fim_rateio = False


def gerar_txt_saida(linha):
    if linha[0] and linha[0] != 'Ev':
        if len(linha) == 10:
            print(f'{linha[0]} | {linha[2]} | {linha[4]} | {linha[5]} | {linha[7]} | {linha[9]}')
        else:
            print(linha)



def pegar_centro_custo(linha):
    if 'Total do Rateio' in linha[0]:
        if linha[0].split('-')[-1].strip() == 'labores':
            return 'Pro-labores'
        if 'Sem Rateio' in linha[0]:
            return None
        return linha[0].split('-')[-1].strip()


def ler_tabela_eventos():
    dic_eventos = {}
    lista_custos = []
    cabecalho = True
    pasta_eventos = load_workbook('eventos.xlsx')
    planilha_eventos = pasta_eventos['planilha_eventos']
    # lista_eventos = [linha for linha in planilha_eventos.values]
    coluna = 4
    for linha in planilha_eventos.values:
        tam_colunas = len(linha)
        if not cabecalho:
            dic_eventos[linha[0]] = {'descricao': linha[1]}
            dic_eventos[linha[0]].update({'hist': linha[2]}) 
            dic_eventos[linha[0]].update({'credito': linha[3]}) 
            if coluna < tam_colunas:
                dic_eventos[linha[0]].update({lista_custos[coluna].get(coluna): linha[coluna]}) 
                coluna += 1
            # print(linha)
        else:
            while(coluna < tam_colunas):
                lista_custos.append({coluna:linha[coluna]})
                coluna += 1
            coluna = 4
            cabecalho = False
    print(lista_custos)

    print(dic_eventos)
    # return lista_eventos



if __name__ == '__main__':
    ler_tabela_eventos()


    with open('Relatorios_Calculo_Relacao_de_Calculo_Rateada.csv') as folha:
        for linha in csv.reader(folha, delimiter=';'):
            if linha:
                centro_de_custo = pegar_centro_custo(linha)
                if centro_de_custo:
                    Rota.centro_custo = True
                    # print(centro_de_custo)
                if 'Ev' in linha and Rota.centro_custo:
                    Rota.evento = True
                if Rota.centro_custo and Rota.evento:
                    # gerar_txt_saida(linha)
                    pass
                if 'Parte RAT' in linha[0]:
                    Rota.centro_custo = False
                    Rota.evento = False

           