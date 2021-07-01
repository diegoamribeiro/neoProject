import win32com.client as win32
import matplotlib.pyplot as plt
from sklearn.metrics import mean_squared_error
from math import sqrt
import pystan
from pandas import DataFrame
from matplotlib import pyplot
from fbprophet import prophet
import pandas as pd
from pandas.plotting import autocorrelation_plot
import warnings
import numpy as np
import pandas.testing as tm
from statsmodels.graphics.tsaplots import plot_acf
from statsmodels.tsa.arima_model import ARIMA
from statsmodels.graphics.tsaplots import plot_pacf
import os


attachment_dir = "anexos\\"
parent_dir = "C:\\temp\\"
path = os.path.join(parent_dir, attachment_dir)

if not os.path.exists(path):
    os.mkdir(path)

path_file = 'Dados_de_empresas_ficticias.csv'

file = pd.read_csv(path_file, sep=';', decimal='.', squeeze=True)
series = pd.read_csv("Dados_de_empresas_ficticias.csv",
                     sep=';', usecols=['Consumo 01/2020', 'Consumo 02/2020', 'Consumo 03/2020', 'Consumo 04/2020',
                     'Consumo 05/2020', 'Consumo 06/2020', 'Consumo 07/2020', 'Consumo 08/2020',
                     'Consumo 09/2020', 'Consumo 10/2020', 'Consumo 11/2020', 'Consumo 12/2020'])

dataframe = pd.DataFrame(file)


consumo_futuro = []

meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

for i in range(0, len(series)):
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

    series['Consumo 01/2020'] = janeiro
    series['Consumo 02/2020'] = fevereiro
    series['Consumo 03/2020'] = marco
    series['Consumo 04/2020'] = abril
    series['Consumo 05/2020'] = maio
    series['Consumo 06/2020'] = junho
    series['Consumo 07/2020'] = julho
    series['Consumo 08/2020'] = agosto
    series['Consumo 09/2020'] = setembro
    series['Consumo 10/2020'] = outubro
    series['Consumo 11/2020'] = novembro
    series['Consumo 12/2020'] = dezembro

split_point = len(series) - 12

dataset = series[0:split_point]
validation = series[split_point:]


series = pd.read_csv('dataset.csv')

new_dataframe = pd.DataFrame({'ds': dataset.index, 'y': dataset.values})
dataframe.head()

model = Prophet()
model.fit(new_dataframe)

futuro = model.make_future_dataframe(periods=12, freq='M')

saida = model.predict(futuro)
saida[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail(7)

validation_df = pd.DataFrame({'ds': validation.index})
saida = model.predict(validation_df)
graphic = model.plot(saida)


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
