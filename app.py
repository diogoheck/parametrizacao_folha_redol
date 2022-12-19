import csv
from openpyxl import load_workbook
import os

EMPRESA = '0001'


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
        valor = linha[4].replace('.', '').replace(',', '').zfill(16)
        # descricao = linha[2]
    elif INSS:
        if codigo_evento == 'PARTE EMPRESA':
            valor = linha[1].replace('.', '').replace(',', '').zfill(16)
        elif codigo_evento == 'PARTE TERCEIROS' or codigo_evento == 'DIRETOR':
            # print('passe aqui')
            valor = linha[3].replace('.', '').replace(',', '').zfill(16)
        else:
            valor = linha[1].replace('.', '').replace(',', '').zfill(16)


    else:
        valor = linha[9].replace('.', '').replace(',', '').zfill(16)
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
    
    if os.path.exists('layout_folha_importacao.txt'):
        os.remove('layout_folha_importacao.txt')

    if os.path.exists('log.txt'):
        os.remove('log.txt')

    centro_de_custo = None
    pegou_data = False
    with open('Relatorios_Calculo_Relacao_de_Calculo_Rateada.csv') as folha:

        for linha in csv.reader(folha, delimiter=';'):
            if linha:
                if 'Período' in linha[0] and not pegou_data:
                    pegou_data = True
                    data = linha[0].split(' ')[3]
                    data_ajuste = data.split('/')[2][2:4] + data.split('/')[1] + data.split('/')[0]
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
            valor = float(linha.strip().split(' ')[-1]) / 100
        
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
    print('\n')
    print('finalizado com sucesso!!')
    print('\n')
    os.system('pause')
    



def converter_string_para_float(valor):
    if '.' in valor and ',' in valor:
        valor = valor.replace('.', '').replace(',', '.')
    elif ',' in valor:
        valor = valor.replace(',', '.')

    return round(float(valor),2)


class Funcionario:
    def __init__(self, codigo, nome) -> None:
        self.codigo = codigo
        self.nome = nome
        self.soma = False
        self.provisao_13 = 0 

    def add_saldo_anterior(self, valor):
        self.provisao_13 += valor

    def add_saldo(self, valor):
        self.provisao_13 += valor

    def add_pago(self, valor):
        self.provisao_13 += valor


def ler_tabela_eventos_provisoes():
    dic_eventos = {}
    lista_custos = []
    cabecalho = True
    pasta_eventos = load_workbook('eventos_provisoes.xlsx')
    planilha_eventos = pasta_eventos['plan_provisoes']
    
    dic_func = {}
    i = 1
    for linha in planilha_eventos.values:
        if i >= 3:
            nome = linha[0]
          
            dic_func[nome] = {}
            dic_func[nome].update({'cc': linha[1]})
            dic_func[nome].update({'sub_c': linha[2]})
            dic_func[nome].update({'prov_13_deb': linha[3]})
            dic_func[nome].update({'prov_13_cred': linha[4]})
            dic_func[nome].update({'prov_13_hist': linha[5]})
            dic_func[nome].update({'fgts_13_deb' : linha[6]})
            dic_func[nome].update({'fgts_13_cred' : linha[7]})
            dic_func[nome].update({'fgts_13_hist' : linha[8]})
            dic_func[nome].update({'fgts_13_hist_bx' : linha[9]})
            dic_func[nome].update({'inss_13_deb' : linha[10]})
            dic_func[nome].update({'inss_13_cred' : linha[11]})
            dic_func[nome].update({'inss_13_hist' : linha[12]})
            dic_func[nome].update({'inss_13_hist_bx' : linha[13]})
            dic_func[nome].update({'prov_ferias_deb' : linha[15]})
            dic_func[nome].update({'prov_ferias_cred' : linha[16]})
            dic_func[nome].update({'prov_ferias_hist' : linha[17]})
            dic_func[nome].update({'fgts_ferias_deb' : linha[18]})
            dic_func[nome].update({'fgts_ferias_cred' : linha[19]})
            dic_func[nome].update({'fgts_ferias_hist' : linha[20]})
            dic_func[nome].update({'fgts_ferias_hist_bx' : linha[21]})
            dic_func[nome].update({'inss_ferias_deb' : linha[22]})
            dic_func[nome].update({'inss_ferias_cred' : linha[23]})
            dic_func[nome].update({'inss_ferias_hist' : linha[24]})
            dic_func[nome].update({'inss_ferias_hist_bx' : linha[25]})
    
        i +=1

    return dic_func
    
            # print(linha)


# def layout_saida_provisoes(dic_func, valor):
#     data = '221130'
#     if float(valor) > 0:
#         print(f'{EMPRESA}{28 * " "}{data}{35 * " "}{conta_credito} {conta_debito}{13 * " "}{historico} {valor}', file=folha)


def lancar_folha_13_salario() :

    with open('centro_de_custo.txt', 'r', encoding='utf-8') as cc:
        arquivo_cc = [centro_custo.replace('\n', '') for centro_custo in cc.readlines()]
    # print(arquivo_cc)
    # ler a folha
    saldo_anterior_boleano = False
    saldo_boleano = False
    pago_boleano = False
    funcionario = False
    total_saldo_anterior = 0.0
    total_saldo = 0.0
    total_pago = 0.0
    dic_func = ler_tabela_eventos_provisoes()

    with open('Relatorios_Funcionarios_Provisoes_Provisao_13o_Grafica.csv', 'r') as folha_13:
        for linha in csv.reader(folha_13, delimiter=';'):
            
    
            if linha:
                # print(linha)
                if len(linha) == 1:
                    if (linha[0].strip().split('-')[0][0:11]).upper() == 'ORGANOGRAMA': 
                        # possivel_centro_custo = linha[0].strip().split('-')[1].strip()
                        # print(linha[0].strip().split('-')[0][0:11])
                        centro_de_custo = linha[0].strip().split('-')[1].strip().upper()
                        # if centro_de_custo in arquivo_cc:
                        #     centro_de_custo = True
                            # print(f'centro de custo =>> {centro_de_custo}')
                try:
                    
                    codigo_funcionario = int(linha[0])
                    nome_funcionario = linha[1]
                    funcionario = True
                    NovoFuncionario = Funcionario(codigo_funcionario, nome_funcionario)
                    # print(f'{nome_funcionario};{centro_de_custo}')
                except:
                    pass

                if funcionario:

                    if linha[0].upper() == 'SALDO ANTERIOR':
                        saldo_anterior = linha[1]
                        saldo_anterior_boleano = True
                        # print(f'saldo anterior => {saldo_anterior}')
                    if linha[0].upper() == 'SALDO':
                        saldo = linha[1]
                        saldo_boleano = True
                        # print(f'saldo => {saldo}')
                    if linha[0].upper() == 'PAGO':
                        pago = linha[1]
                        pago_boleano = True
                        # print(f'pago => {pago}')
                    # if centro_custo and funcionario:
                if saldo_anterior_boleano and saldo_boleano and pago_boleano and funcionario:
                    funcionario = False
                    saldo_anterior_boleano = False
                    saldo_boleano = False
                    pago_boleano = False
                    saldo = converter_string_para_float(saldo)
                    saldo_anterior = converter_string_para_float(saldo_anterior)
                    pago = converter_string_para_float(pago)
                    provisao = float(saldo) + float(pago) - float(saldo_anterior)
                    # print(saldo_anterior, saldo, pago)
                    total_saldo_anterior += saldo_anterior
                    total_saldo += saldo
                    total_pago += pago
                    if dic_func.get(nome_funcionario, 'nao encontrado'):
                        # layout_saida_provisoes(dic_func, provisao)
                        if provisao > 0:
                            prov_deb = dic_func.get(nome_funcionario)['prov_13_deb']
                            prov_cred = dic_func.get(nome_funcionario)['prov_13_cred']
                            print(nome_funcionario, prov_deb, prov_cred, round(provisao, 2))
                        elif provisao < 0:
                            provisao *= -1
                            prov_cred = dic_func.get(nome_funcionario)['prov_13_deb']
                            prov_deb = dic_func.get(nome_funcionario)['prov_13_cred']
                            print(nome_funcionario, prov_deb, prov_cred, round(provisao, 2))

                        else:
                            pass
                        # print(f'centro de custo {centro_de_custo} funcionario {nome_funcionario} provisao_lancar {round(provisao, 2)}')
                    else:
                        print('nao encontrado')
        print(f'total saldo anterior {round(total_saldo_anterior, 2)}')
        print(f'total saldo {round(total_saldo, 2)}')
        print(f'total pago {round(total_pago, 2)}')
        print(f'total provisao a lançar {round(total_saldo + total_pago - total_saldo_anterior, 2)}')
    

if __name__ == '__main__':
    lancar_folha_pagamento()
    # lancar_folha_13_salario()

