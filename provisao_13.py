import csv
from openpyxl import load_workbook
import os

EMPRESA = '0001'

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


def gerar_lancamentos_13_salario(conta_debito, conta_credito, valor, historico, data):
    conta_debito = str(conta_debito).replace('-', '').zfill(7) 
    conta_credito = str(conta_credito).replace('-', '').zfill(19)
    valor = str(valor)
    valor = valor.replace('.', '').replace(',', '').zfill(15) + '1'
    historico = str(historico).zfill(4)
    with open('layout_folha_importacao.txt', 'a') as folha:
        if float(valor) > 0:
            print(f'{EMPRESA}{28 * " "}{data}{35 * " "}{conta_debito} {conta_credito}{13 * " "}{historico} {valor}', file=folha)


def lancar_folha_13_salario(data) :

    # with open('centro_de_custo.txt', 'r', encoding='utf-8') as cc:
    #     arquivo_cc = [centro_custo.replace('\n', '') for centro_custo in cc.readlines()]
    # print(arquivo_cc)
    # ler a folha
    saldo_anterior_boleano = False
    saldo_boleano = False
    pago_boleano = False
    funcionario = False
    total_saldo_anterior = 0.0
    total_saldo_anterior_INSS = 0.0
    total_saldo_anterior_FGTS = 0.0
    total_saldo = 0.0
    total_saldo_INSS = 0.0
    total_saldo_FGTS = 0.0
    total_pago = 0.0
    total_pago_INSS = 0.0
    total_pago_FGTS = 0.0
    dic_func = ler_tabela_eventos_provisoes()
    with open('log_provisoes.txt', 'a') as log:
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
                            saldo_anterior_INSS = linha[2]
                            saldo_anterior_FGTS = linha[3]
                            saldo_anterior_boleano = True
                            # print(f'saldo anterior => {saldo_anterior}')
                        if linha[0].upper() == 'SALDO':
                            saldo = linha[1]
                            saldo_INSS = linha[2]
                            saldo_FGTS = linha[3]
                            saldo_boleano = True
                            # print(f'saldo => {saldo}')
                        if linha[0].upper() == 'PAGO':
                            pago = linha[1]
                            pago_INSS = linha[2]
                            pago_FGTS = linha[3]
                            pago_boleano = True
                            # print(f'pago => {pago}')
                        # if centro_custo and funcionario:
                    if saldo_anterior_boleano and saldo_boleano and pago_boleano and funcionario:
                        funcionario = False
                        saldo_anterior_boleano = False
                        saldo_boleano = False
                        pago_boleano = False
                        # provisao 13
                        saldo = converter_string_para_float(saldo)
                        saldo_anterior = converter_string_para_float(saldo_anterior)
                        pago = converter_string_para_float(pago)
                        provisao = float(saldo) + float(pago) - float(saldo_anterior)
                        # provisao 13 INSS
                        saldo_INSS = converter_string_para_float(saldo_INSS)
                        saldo_anterior_INSS = converter_string_para_float(saldo_anterior_INSS)
                        pago_INSS = converter_string_para_float(pago_INSS)
                        provisao_13_INSS = float(saldo_INSS) + float(pago_INSS) - float(saldo_anterior_INSS)
                        # provisao 13 FGTS
                        saldo_FGTS = converter_string_para_float(saldo_FGTS)
                        saldo_anterior_FGTS = converter_string_para_float(saldo_anterior_FGTS)
                        pago_FGTS = converter_string_para_float(pago_FGTS)
                        provisao_13_FGTS = float(saldo_FGTS) + float(pago_FGTS) - float(saldo_anterior_FGTS)
                        # print(saldo_anterior, saldo, pago)
                        # totais provisao 13
                        total_saldo_anterior += saldo_anterior
                        total_saldo += saldo
                        total_pago += pago
                        # totais provisao 13 iNSS
                        total_saldo_anterior_INSS += saldo_anterior_INSS
                        total_saldo_INSS += saldo_INSS
                        total_pago_INSS += pago_INSS
                        if dic_func.get(nome_funcionario):
                            # layout_saida_provisoes(dic_func, provisao)
                            if provisao > 0:
                                # print(nome_funcionario)
                                prov_deb = dic_func.get(nome_funcionario)['prov_13_deb']
                                prov_cred = dic_func.get(nome_funcionario)['prov_13_cred']
                                prov_hist = dic_func.get(nome_funcionario)['prov_13_hist']
                                # print(nome_funcionario, prov_deb, prov_cred, round(provisao, 2), prov_hist)
                                gerar_lancamentos_13_salario(prov_deb, prov_cred, round(provisao, 2), prov_hist, data)
                            elif provisao < 0:
                                provisao *= -1
                                prov_cred = dic_func.get(nome_funcionario)['prov_13_deb']
                                prov_deb = dic_func.get(nome_funcionario)['prov_13_cred']
                                prov_hist = dic_func.get(nome_funcionario)['prov_13_hist']
                                # print(nome_funcionario, prov_deb, prov_cred, round(provisao, 2))
                                gerar_lancamentos_13_salario(prov_deb, prov_cred, round(provisao, 2), prov_hist, data)

                            if provisao_13_INSS > 0:
                                prov_deb_INSS = dic_func.get(nome_funcionario)['inss_13_deb']
                                prov_cred_INSS = dic_func.get(nome_funcionario)['inss_13_cred']
                                prov_hist_INSS = dic_func.get(nome_funcionario)['inss_13_hist']
                                gerar_lancamentos_13_salario(prov_deb_INSS, prov_cred_INSS, round(provisao_13_INSS, 2), prov_hist_INSS, data)
                            elif provisao_13_INSS < 0:
                                provisao_13_INSS *= -1
                                prov_cred_INSS = dic_func.get(nome_funcionario)['inss_13_deb']
                                prov_deb_INSS = dic_func.get(nome_funcionario)['inss_13_cred']
                                prov_hist_INSS = dic_func.get(nome_funcionario)['inss_13_hist']
                                gerar_lancamentos_13_salario(prov_deb_INSS, prov_cred_INSS, round(provisao_13_INSS, 2), prov_hist_INSS, data)
                            if pago_INSS > 0:
                                prov_cred_INSS = dic_func.get(nome_funcionario)['inss_13_deb']
                                prov_deb_INSS = dic_func.get(nome_funcionario)['inss_13_cred']
                                prov_hist_INSS = dic_func.get(nome_funcionario)['inss_13_hist_bx']
                                gerar_lancamentos_13_salario(prov_deb_INSS, prov_cred_INSS, round(pago_INSS, 2), prov_hist_INSS, data)

                            if provisao_13_FGTS > 0:
                                prov_deb_FGTS = dic_func.get(nome_funcionario)['fgts_13_deb']
                                prov_cred_FGTS = dic_func.get(nome_funcionario)['fgts_13_cred']
                                prov_hist_FGTS = dic_func.get(nome_funcionario)['fgts_13_hist']
                                gerar_lancamentos_13_salario(prov_deb_FGTS, prov_cred_FGTS, round(provisao_13_FGTS, 2), prov_hist_FGTS, data)
                            elif provisao_13_FGTS < 0:
                                provisao_13_FGTS *= -1
                                prov_cred_FGTS = dic_func.get(nome_funcionario)['fgts_13_deb']
                                prov_deb_FGTS = dic_func.get(nome_funcionario)['fgts_13_cred']
                                prov_hist_FGTS = dic_func.get(nome_funcionario)['fgts_13_hist']
                                gerar_lancamentos_13_salario(prov_deb_FGTS, prov_cred_FGTS, round(provisao_13_FGTS, 2), prov_hist_FGTS, data)
                            if pago_FGTS > 0:
                                prov_cred_FGTS = dic_func.get(nome_funcionario)['fgts_13_deb']
                                prov_deb_FGTS = dic_func.get(nome_funcionario)['fgts_13_cred']
                                prov_hist_FGTS = dic_func.get(nome_funcionario)['fgts_13_hist_bx']
                                gerar_lancamentos_13_salario(prov_deb_FGTS, prov_cred_FGTS, round(pago_FGTS, 2), prov_hist_FGTS, data)

                            else:
                                pass
                            # print(f'centro de custo {centro_de_custo} funcionario {nome_funcionario} provisao_lancar {round(provisao, 2)}')
                        else:
                            print(f'{nome_funcionario} nao encontrado {centro_de_custo}', file=log)
                            
            # print(f'total saldo anterior {round(total_saldo_anterior, 2)}')
            # print(f'total saldo {round(total_saldo, 2)}')
            # print(f'total pago {round(total_pago, 2)}')
            # print(f'total provisao a lançar {round(total_saldo + total_pago - total_saldo_anterior, 2)}')