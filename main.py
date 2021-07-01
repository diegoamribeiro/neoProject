import win32com.client as win32
import matplotlib.pyplot as plt
from pandas import DataFrame
from matplotlib import pyplot
import pandas as pd
from pandas.plotting import autocorrelation_plot
import warnings
import numpy as np
import pandas.testing as tm
from statsmodels.graphics.tsaplots import plot_acf
from statsmodels.tsa.arima_model import ARIMA
from statsmodels.graphics.tsaplots import plot_pacf

import os

warnings.filterwarnings("ignore")


def difference(dataset, interval=1):
    diff = list()
    for i in range(interval, len(dataset)):
        value = dataset[i] - dataset[i - interval]
        diff.append(value)
    return diff


def inverse_difference(historic, prediction, interval=1):
    return prediction + historic[-interval]


attachment_dir = "anexos\\"
parent_dir = "C:\\temp\\"
path = os.path.join(parent_dir, attachment_dir)

if not os.path.exists(path):
    os.mkdir(path)

path_file = 'Dados_de_empresas_ficticias.csv'

file = pd.read_csv(path_file, sep=';', decimal='.', squeeze=True)
file_2 = pd.read_csv("Dados_de_empresas_ficticias.csv",
                     sep=';', usecols=['Consumo 01/2020', 'Consumo 02/2020', 'Consumo 03/2020', 'Consumo 04/2020',
                     'Consumo 05/2020', 'Consumo 06/2020', 'Consumo 07/2020', 'Consumo 08/2020',
                     'Consumo 09/2020', 'Consumo 10/2020', 'Consumo 11/2020', 'Consumo 12/2020'])

dataframe = pd.DataFrame(file)


consumo_futuro = []
for i in range(0, len(file_2)):
    janeiro = float(dataframe['Consumo 01/2020'][i].replace(',', '.'))
    fevereiro = float(dataframe['Consumo 02/2020'][i].replace(',', '.'))
    marco = float(dataframe['Consumo 03/2020'][i].replace(',', '.'))
    abril = float(dataframe['Consumo 04/2020'][i].replace(',', '.'))
    maio = float(dataframe['Consumo 05/2020'][i].replace(',', '.'))
    junho = float(dataframe['Consumo 06/2020'][i].replace(',', '.'))
    julho = float(dataframe['Consumo 07/2020'][i].replace(',', '.'))
    agosto = float(dataframe['Consumo 08/2020'][i].replace(',', '.'))
    setembro = float(dataframe['Consumo 09/2020'][i].replace(',', '.'))
    outubro = float(dataframe['Consumo 10/2020'][i].replace(',', '.'))
    novembro = float(dataframe['Consumo 11/2020'][i].replace(',', '.'))
    dezembro = float(dataframe['Consumo 12/2020'][i].replace(',', '.'))

    file_2['Consumo 01/2020'] = janeiro
    file_2['Consumo 02/2020'] = fevereiro
    file_2['Consumo 03/2020'] = marco
    file_2['Consumo 04/2020'] = abril
    file_2['Consumo 05/2020'] = maio
    file_2['Consumo 06/2020'] = junho
    file_2['Consumo 07/2020'] = julho
    file_2['Consumo 08/2020'] = agosto
    file_2['Consumo 09/2020'] = setembro
    file_2['Consumo 10/2020'] = outubro
    file_2['Consumo 11/2020'] = novembro
    file_2['Consumo 12/2020'] = dezembro


#consumo_futuro = [janeiro, fevereiro, marco, abril, maio, junho, julho, agosto, setembro,
#                   outubro, novembro, dezembro]

# modelo = ARIMA(consumo_futuro, order=(0, 1, 1))
#
# modelo_fit = modelo.fit()
#
# # print(modelo_fit.summary())
#
# residuals = DataFrame(modelo_fit.resid)
# residuals.plot()
# residuals.plot(kind='kde')
# pyplot.show()

# print(residuals.describe())

X = file_2.values

# print(X)
X = X.astype('float32')

# Divide a fonte de dados
size = int(len(X) * 0.50)

# Separa os dados de treino e teste
train = X[0:size]
test = X[size:]

# controle de dados
history = [x for x in train]

# cria lista de previsões
predictions = list()

for t in range(len(test)):
    meses_do_ano = 12
    diff = difference(history, meses_do_ano)
    modelo = ARIMA(diff, order=(1, 1, 1))
    modelo_fit = modelo.fit(trend='nc', disp=0)

# plot_acf(consumo_futuro, lags=11)
# pyplot.show()
# file_2.info()


outlook = win32.Dispatch('outlook.application')


def send_email(destino, anexo):
    custom_email = outlook.CreateItem(0)
    custom_email.To = destino
    custom_email.Subject = 'Previsão de carga - Neoenergia'
    custom_email.HTMLBody = """
        <style>
        h1 {
          color: #427314;
          text-indent: 18px;
          text-transform: uppercase;
          font-family: Arial,sans-serif;
          font-weight: bold;
          font-size: 15px
        }
</style>
        <p>Caro cliente,</p>
        <p>segue em anexo a previsão de carga.</p>
        </br>
        </br>
        </br>

        <p>Atenciosamente,</p>
        
        <h1>Neoenergia</h1>
        """
    custom_email.Attachments.Add(anexo)
    custom_email.Send()


# send_email()


meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

for i in range(0, len(dataframe)):
    empresa = dataframe['Empresa'][i]
    email = dataframe["Email do responsavel"][i]

    janeiro = float(dataframe['Consumo 01/2020'][i].replace(',', '.'))
    fevereiro = float(dataframe['Consumo 02/2020'][i].replace(',', '.'))
    marco = float(dataframe['Consumo 03/2020'][i].replace(',', '.'))
    abril = float(dataframe['Consumo 04/2020'][i].replace(',', '.'))
    maio = float(dataframe['Consumo 05/2020'][i].replace(',', '.'))
    junho = float(dataframe['Consumo 06/2020'][i].replace(',', '.'))
    julho = float(dataframe['Consumo 07/2020'][i].replace(',', '.'))
    agosto = float(dataframe['Consumo 08/2020'][i].replace(',', '.'))
    setembro = float(dataframe['Consumo 09/2020'][i].replace(',', '.'))
    outubro = float(dataframe['Consumo 10/2020'][i].replace(',', '.'))
    novembro = float(dataframe['Consumo 11/2020'][i].replace(',', '.'))
    dezembro = float(dataframe['Consumo 12/2020'][i].replace(',', '.'))

    consumo = [janeiro, fevereiro, marco, abril, maio, junho, julho, agosto, setembro,
               outubro, novembro, dezembro]

    plt.plot(meses, consumo, c='#427314')
    plt.title(f"{empresa} - Ano 2020")

    #
    #  plt.grid()
    # fig = plt.savefig(f"{path}/{empresa}.png")
    # plt.show()
    # send_email(email, f"{path}/{empresa}.png")
    # print(send_email(email, f"{path}/{empresa}.png"))

#     fig, ax = plt.subplots()
#     x = np.arange(len(meses))
#     width = 0.35
#     rects1 = ax.bar(x - width / 2, meses, width, label='Men')
#     rects2 = ax.bar(x + width / 2, consumo, width, label='Women')
#     plt.bar(meses, consumo)
#     plt.title(f"{empresa} - Ano 2020")

# if __name__ == '__main__':
#     print_hi('PyCharm')
# meses_file = file_2[['Consumo 01/2020', 'Consumo 02/2020', 'Consumo 03/2020', 'Consumo 04/2020',
#                      'Consumo 05/2020', 'Consumo 06/2020', 'Consumo 07/2020', 'Consumo 08/2020',
#                      'Consumo 09/2020', 'Consumo 10/2020', 'Consumo 11/2020', 'Consumo 12/2020']]
