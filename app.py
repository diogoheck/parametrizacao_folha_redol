import csv
from openpyxl import load_workbook
import os
from provisao_13 import lancar_folha_13_salario
from provisao_ferias import lancar_folha_ferias_salario
EMPRESA = '0001'


class DataArquivo:
    data = '01011970'

class Rota:
    centro_custo = False
    evento = False
    fim_rateio = False

#  centro_de_custo = 'Pro-labores'

def auxiliar_folha(codigo_evento, folha, valor, historico, data, centro_de_custo, tabela_eventos):

    if tabela_eventos.get(codigo_evento).get('tipo') == 'P':
        conta_debito = str(tabela_eventos.get(codigo_evento)[centro_de_custo]).replace('-', '').zfill(7) 
        conta_credito = str(tabela_eventos.get(codigo_evento)["credito"]).replace('-', '').zfill(19)
        if float(valor) > 0:
            print(f'{EMPRESA}{28 * " "}{data}{35 * " "}{conta_debito} {conta_credito}{13 * " "}{historico} {valor}', file=folha)
    elif tabela_eventos.get(codigo_evento).get('tipo') == 'D':
        conta_debito = str(tabela_eventos.get(codigo_evento)[centro_de_custo]).replace('-', '').zfill(19) 
        conta_credito = str(tabela_eventos.get(codigo_evento)["credito"]).replace('-', '')
        if conta_credito == '20370' and centro_de_custo == 'Pro-labores':
            conta_credito = '20420'.zfill(7)
        else:
            conta_credito = str(tabela_eventos.get(codigo_evento)["credito"]).replace('-', '').zfill(7) 
            
        if float(valor) > 0:
            print(f'{EMPRESA}{28 * " "}{data}{35 * " "}{conta_credito} {conta_debito}{13 * " "}{historico} {valor}', file=folha)
        


def layout_folha_sistema_redol(tabela_eventos, codigo_evento, folha, centro_de_custo, linha, provento, data, INSS):
    historico = str(tabela_eventos.get(codigo_evento)["hist"]).zfill(4)
    if provento:
        valor = linha[4].replace('.', '').replace(',', '').zfill(15) + '1'
        # descricao = linha[2]
    elif INSS:
        if codigo_evento == 'PARTE EMPRESA':
            valor = linha[1].replace('.', '').replace(',', '').zfill(15) + '1'
        elif codigo_evento == 'PARTE TERCEIROS' or codigo_evento == 'DIRETOR':
            # print('passe aqui')
            valor = linha[3].replace('.', '').replace(',', '').zfill(15) + '1'
        else:
            valor = linha[1].replace('.', '').replace(',', '').zfill(15) + '1'


    else:
        valor = linha[9].replace('.', '').replace(',', '').zfill(15) + '1'
        # descricao = linha[7]
    auxiliar_folha(codigo_evento, folha, valor, historico, data, centro_de_custo, tabela_eventos)


    
def lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS):
    if codigo_evento:
        if tabela_eventos.get(codigo_evento):
            
            if tabela_eventos.get(codigo_evento).get(centro_de_custo, 'NAO ENCONTRADO'):
                    layout_folha_sistema_redol(tabela_eventos, codigo_evento, folha, centro_de_custo, linha, provento, data, INSS)

            elif tabela_eventos.get(codigo_evento).get(centro_de_custo, 'NAO ENCONTRADO') == 'NAO ENCONTRADO':
                print(f'centro de custos {centro_de_custo} nao encontrado ', file=log)
                    
        else:
            # print(linha)
            # print(codigo_evento)
            print(f'evento {codigo_evento} nao encontrado referente centro custo {centro_de_custo}', file=log)
            
    
    

def gerar_txt_saida(linha, tabela_eventos, centro_de_custo, data):
    provento = True
    INSS = False
    with open('log.txt', 'a') as log:
        with open('layout_folha_importacao.txt', 'a', encoding='utf-8') as folha:
            provento = False
            INSS = False
            if linha[0] != 'Ev':
                # print(linha, len(linha))
            
                
                # proventos e descontos
                if len(linha) > 3 and len(linha) <= 10:
                    try:
                        codigo_evento = int(linha[0])
                        # print(linha, len(linha))

                        if len(linha) == 10:
                            # lanca provento
                            provento = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)
                            # lanca desconto
                            if linha[5]:
                                codigo_evento = int(linha[5])
                                provento = False
                                lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)
                        elif len(linha) == 5:
                            provento = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)
                         

                    except:
                        
                        if len(linha) == 10:
                            try:
                                codigo_evento = int(linha[5])
                                provento = False
                                lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)
                            except:
                                pass
                        
                        # print(linha)
                        if linha[0].replace(':', '').strip().upper() == 'PARTE EMPRESA':
                            codigo_evento = linha[0].replace(':', '').strip().upper()
                            # print(codigo_evento, linha[1], centro_de_custo)
                            INSS = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)

                        if linha[0].replace(':', '').strip().upper() == 'ENTIDADE FINANCEIRA':
                            codigo_evento = linha[2].replace(':', '').strip().upper()
                            # print(codigo_evento, linha[3], centro_de_custo)
                            INSS = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)
                        if linha[0].replace(':', '').strip().upper() == 'PARTE RAT + ACRÉS. FAP':
                            codigo_evento = linha[0].replace(':', '').strip().upper()
                            # print(codigo_evento, linha[1], centro_de_custo)
                            INSS = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)
                        if linha[0].replace(':', '').strip().upper() == 'SEGURADOS' and centro_de_custo == 'Pro-labores':
                            codigo_evento = linha[2].replace(':', '').strip().upper()
                            # print(codigo_evento, linha[1], centro_de_custo)
                            INSS = True
                            lancar_folha(codigo_evento, log, tabela_eventos, centro_de_custo, linha, folha, provento, data, INSS)
                        
                        

                

                        


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



def lancar_folha_pagamento():

    tabela_eventos = ler_tabela_eventos()

    # print(tabela_eventos)
    
    

    centro_de_custo = None
    pegou_data = False
    with open('Relatorios_Calculo_Relacao_de_Calculo_Rateada.csv') as folha:

        for linha in csv.reader(folha, delimiter=';'):
            if linha:
                if 'Período' in linha[0] and not pegou_data:
                    pegou_data = True
                    data = linha[0].split(' ')[3]
                    data_ajuste = data.split('/')[2][2:4] + data.split('/')[1] + data.split('/')[0]
                    DataArquivo.data = data_ajuste
                    # print(data_ajuste)

                if 'Total do Rateio' in linha[0]:
                    if linha[0].split('-')[-1].strip() == 'labores':
                        centro_de_custo = 'Pro-labores'
                        Rota.centro_custo = True
                    elif 'Sem Rateio' in linha[0]:
                        centro_de_custo = None
                    else:
                        centro_de_custo = linha[0].split('-')[-1].strip()
                        Rota.centro_custo = True
                
                if 'Ev' in linha and Rota.centro_custo:
                    Rota.evento = True
                if Rota.centro_custo and Rota.evento:
                    gerar_txt_saida(linha, tabela_eventos, centro_de_custo, data_ajuste)
                if 'Parte RAT' in linha[0]:
                    Rota.centro_custo = False
                    Rota.evento = False


    # eliminar repeticoes no arquivo de log
    if os.path.exists('log.txt'):
        conjunto = tuple()
        with open('log.txt', 'r') as log:
            conjunto = set(log.readlines())
            
    os.remove('log.txt')
    # print(conjunto)

    with open('log.txt', 'a') as log:
        for conj in conjunto:
            print(f'{conj.strip()}', file=log)

    total_proventos = 0
    total_descontos = 0
    total_FGTS = 0
    total_inss = 0
    total_IRRF = 0
    total_prolabore = 0
    # conferencia da folha
    with open('layout_folha_importacao.txt', 'r') as folha:
        for linha in folha.readlines():
            # print(linha.strip())
            credito = linha.strip().split(' ')[-15][14::]
            debito = linha.strip().split(' ')[-16][2::]
            valor = float(linha.strip().split(' ')[-1][0:15]) / 100

        
            if credito == '20370' or credito == '20420':
                total_proventos += valor
            if debito == '20370' or debito == '20420':
                total_descontos += valor
            if credito == '20440':
                total_FGTS += valor
            if credito == '20430':
                total_inss += valor
            if debito == '20430':
                total_inss -= valor
            if credito == '20490':
                total_IRRF += valor
            if debito == '20490':
                total_IRRF -= valor
            if credito == '20420':
                total_prolabore += valor
            if debito == '20420':
                total_prolabore -= valor



        with open('memoria_calculo.txt', 'w') as memoria:
        
            print(f'Total dos proventos: {round(total_proventos, 2)}', file=memoria)
            print(f'Total de descontos: {round(total_descontos, 2)}', file=memoria)
            print(f'FGTS a pagar: {round(total_FGTS, 2)}',file=memoria)
            print(f'INSS a pagar: {round(total_inss, 2)}', file=memoria)
            print(f'IRRF a pagar: {round(total_IRRF, 2)}', file=memoria)
            print(f'Pro-labore a pagar: {round(total_prolabore, 2)}', file=memoria)
 
    



def converter_string_para_float(valor):
    if '.' in valor and ',' in valor:
        valor = valor.replace('.', '').replace(',', '.')
    elif ',' in valor:
        valor = valor.replace(',', '.')

    return round(float(valor),2)






    
            # print(linha)


# def layout_saida_provisoes(dic_func, valor):
#     data = '221130'
#     if float(valor) > 0:
#         print(f'{EMPRESA}{28 * " "}{data}{35 * " "}{conta_credito} {conta_debito}{13 * " "}{historico} {valor}', file=folha)
def log_erro(msg):
    with open('log_provisoes.txt', 'r') as log:
        print(msg, file=log)





    

if __name__ == '__main__':
    if os.path.exists('layout_folha_importacao.txt'):
        os.remove('layout_folha_importacao.txt')

    if os.path.exists('log.txt'):
        os.remove('log.txt')
    lancar_folha_pagamento()
    lancar_folha_13_salario(DataArquivo.data)
    lancar_folha_ferias_salario(DataArquivo.data)

    print('\n')
    print('finalizado com sucesso!!')
    print('\n')
    os.system('pause')