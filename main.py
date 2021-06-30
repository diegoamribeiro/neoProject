import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import ast

import locale

dados = pd.read_csv('C:/Users/COELBA/PycharmProjects/neoProject/Dados_de_empresas_ficticias.csv', sep=';', decimal='.')

# print(dados.head())
# meses = {'Janeiro': dados['Consumo 01/2020'], 'Fevereiro': dados['Consumo 02/2020']}
# print(meses.keys())

meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
values = [96.3, 196.2, 390.8, 235, 155.2, 317.4, 158.7, 78.7, 383.775, 20.036, 99.322, 35.6]

for i in range(0, len(dados)):
    empresa = dados['Empresa'][i]
    email = dados['Email do responsavel'][i]

    janeiro = dados['Consumo 01/2020'][i].replace()
    fevereiro = dados['Consumo 02/2020'][i]
    marco = dados['Consumo 03/2020'][i]
    abril = dados['Consumo 04/2020'][i]
    maio = dados['Consumo 05/2020'][i]
    junho = dados['Consumo 06/2020'][i]
    julho = dados['Consumo 07/2020'][i]
    agosto = dados['Consumo 08/2020'][i]
    setembro = dados['Consumo 09/2020'][i]
    outubro = dados['Consumo 10/2020'][i]
    novembro = dados['Consumo 11/2020'][i]
    dezembro = dados['Consumo 12/2020'][i]

    consumo = [janeiro.str.replace(',', '.'), fevereiro, marco, abril, maio, junho, julho, agosto, setembro,
               outubro, novembro, dezembro]

    # type(dezembro)
    print(dezembro)
    # mjaneiro = {'Jan': consumo[0]}
    # mfevereiro = {'Fev': consumo[0]}
    # mmarco = {'Mar': consumo[0]}

    # print(meses, consumo, email)
    #
    # plt.bar(meses, values)
    # plt.show()

    # xposition = np.arange(meses[0],meses, 1)
    # print(xposition)
    # print(meses)

    # plt.xticks(meses, consumo)

    # print(empresa)
    # print(ano_2020)
    # print(email)

    # plt.bar(janeiro, ano_2020)

# plt.scatter(x, y)
# plt.show()
#
# if __name__ == '__main__':
#     print_hi('PyCharm')
