import csv
from openpyxl import load_workbook
import os

class Rota:
    centro_custo = False
    evento = False
    fim_rateio = False


def auxiliar_folha(codigo_evento, conta_debito, conta_credito, folha, descricao, valor):
    if tabela_eventos.get(codigo_evento).get('tipo') == 'P':
        print(f'{conta_debito}|{conta_credito}|{descricao}|{valor}|', file=folha)
    else:
        print(f'{conta_credito}|{conta_debito}{descricao}|{valor}|', file=folha)


def layout_folha_sistema_redol(tabela_eventos, codigo_evento, folha, centro_de_custo, linha, provento):
    conta_debito = tabela_eventos.get(codigo_evento)[centro_de_custo] 
    conta_credito = tabela_eventos.get(codigo_evento)["credito"]
    if provento:
        valor = linha[4]
        descricao = linha[2]
    else:
        valor = linha[9]
        descricao = linha[7]
    auxiliar_folha(codigo_evento, conta_debito, conta_credito, folha, descricao, valor)


    
def lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento):
    if codigo_evento:
        if tabela_eventos.get(codigo_evento):
            
            if tabela_eventos.get(codigo_evento).get(centro_de_custo, 'NAO ENCONTRADO'):
                    layout_folha_sistema_redol(tabela_eventos, codigo_evento, folha, centro_de_custo, linha, provento)

            elif tabela_eventos.get(codigo_evento).get(centro_de_custo, 'NAO ENCONTRADO') == 'NAO ENCONTRADO':
                print(f'centro de custos {centro_de_custo} nao encontrado ', file=log)    
        else:
            print(f'evento {codigo_evento} nao encontrado', file=log)


def gerar_txt_saida(linha, tabela_eventos, centro_de_custo):
    provento = True

    with open('log.txt', 'a') as log:
        with open('layout_folha_importacao.txt', 'a', encoding='utf-8') as folha:
            if linha[0] and linha[0] != 'Ev':
                
                # proventos e descontos
                if len(linha) > 3 and len(linha) <= 10:
                    try:
                        codigo_evento = int(linha[0])
                        print(linha, len(linha))

                        if len(linha) == 10:
                            # lanca provento
                            provento = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento)
                            # lanca desconto
                            codigo_evento = linha[5]
                            provento = False
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento)
                        elif len(linha) == 5:
                            provento = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento)
                         

                    except:
                        pass


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
    
    coluna = 5
    for linha in planilha_eventos.values:
        tam_colunas = len(linha)
        if cabecalho:
            lista_custos.append('')
            lista_custos.append('')
            lista_custos.append('')
            lista_custos.append('')
            lista_custos.append('')
            while(coluna < tam_colunas):
                lista_custos.append({coluna:linha[coluna]})
                coluna += 1
            coluna = 5
            cabecalho = False
            
        else:
            coluna = 5
            dic_eventos[linha[0]] = {'tipo': linha[1]}
            dic_eventos[linha[0]].update({'descricao': linha[2]})
            dic_eventos[linha[0]].update({'hist': linha[3]}) 
            dic_eventos[linha[0]].update({'credito': linha[4]}) 
            while coluna < tam_colunas:
                dic_eventos[linha[0]].update({lista_custos[coluna].get(coluna): linha[coluna]}) 
                coluna += 1
    
    
    return dic_eventos



if __name__ == '__main__':
    tabela_eventos = ler_tabela_eventos()

    # print(tabela_eventos)
    
    if os.path.exists('layout_folha_importacao.txt'):
        os.remove('layout_folha_importacao.txt')

    if os.path.exists('log.txt'):
        os.remove('log.txt')

    centro_de_custo = None
    with open('Relatorios_Calculo_Relacao_de_Calculo_Rateada.csv') as folha:

        for linha in csv.reader(folha, delimiter=';'):
            if linha:
                if 'Total do Rateio' in linha[0]:
                    if linha[0].split('-')[-1].strip() == 'labores':
                        centro_de_custo = 'Pro-labores'
                    elif 'Sem Rateio' in linha[0]:
                        centro_de_custo = None
                    else:
                        centro_de_custo = linha[0].split('-')[-1].strip()
                if centro_de_custo:
                    Rota.centro_custo = True
                if 'Ev' in linha and Rota.centro_custo:
                    Rota.evento = True
                if Rota.centro_custo and Rota.evento:
                    gerar_txt_saida(linha, tabela_eventos, centro_de_custo)
                if 'Parte RAT' in linha[0]:
                    Rota.centro_custo = False
                    Rota.evento = False

           