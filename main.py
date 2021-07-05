import win32com.client as win32
import matplotlib.pyplot as plt
import numpy
from sklearn import tree
from sklearn.metrics import mean_squared_error
from sklearn.feature_selection import SelectKBest
from sklearn.model_selection import GridSearchCV
from sklearn.neural_network import MLPRegressor
from sklearn.preprocessing import MinMaxScaler
from sklearn import datasets, linear_model
from sklearn.metrics import mean_squared_error, r2_score
from math import sqrt
import unicodedata

from pandas import DataFrame

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
parent_dir = "C:\\"
path = os.path.join(parent_dir, attachment_dir)

if not os.path.exists(path):
    os.mkdir(path)

path_file = 'Dados_de_empresas_ficticias.csv'

file = pd.read_csv(path_file, sep=';', decimal='.', squeeze=True)

# def remove_non_ascii_normalized(string: str) -> str:
#     normalized = unicodedata.normalize('NFD', string)
#     return normalized.encode('ascii', 'ignore').decode('utf8').casefold()


# palavra = 'Email do responsável'
#
# res_str = palavra[:5] + palavra[:0]
# print(res_str)


dataframe = pd.DataFrame(file)

# print('valor: ', unicodedata.name('á'))
# dataframe.info()


# new = dataframe.loc[0:101, 'Consumo 01/2020': 'Consumo 12/2020']
# consumo = new
#
# print(consumo)


for h in range(0, len(dataframe)):
    dataframe['Consumo 01/2020'][h] = dataframe['Consumo 01/2020'][h].replace(',', '.')
    dataframe['Consumo 02/2020'][h] = dataframe['Consumo 02/2020'][h].replace(',', '.')
    dataframe['Consumo 03/2020'][h] = dataframe['Consumo 03/2020'][h].replace(',', '.')
    dataframe['Consumo 04/2020'][h] = dataframe['Consumo 04/2020'][h].replace(',', '.')
    dataframe['Consumo 05/2020'][h] = dataframe['Consumo 05/2020'][h].replace(',', '.')
    dataframe['Consumo 06/2020'][h] = dataframe['Consumo 06/2020'][h].replace(',', '.')
    dataframe['Consumo 07/2020'][h] = dataframe['Consumo 07/2020'][h].replace(',', '.')
    dataframe['Consumo 08/2020'][h] = dataframe['Consumo 08/2020'][h].replace(',', '.')
    dataframe['Consumo 09/2020'][h] = dataframe['Consumo 09/2020'][h].replace(',', '.')
    dataframe['Consumo 10/2020'][h] = dataframe['Consumo 10/2020'][h].replace(',', '.')
    dataframe['Consumo 11/2020'][h] = dataframe['Consumo 11/2020'][h].replace(',', '.')
    dataframe['Consumo 12/2020'][h] = dataframe['Consumo 12/2020'][h].replace(',', '.')
    dataframe['Consumo 01/2020'][h] = float(dataframe['Consumo 01/2020'][h])
    dataframe['Consumo 02/2020'][h] = float(dataframe['Consumo 02/2020'][h])
    dataframe['Consumo 03/2020'][h] = float(dataframe['Consumo 03/2020'][h])
    dataframe['Consumo 04/2020'][h] = float(dataframe['Consumo 04/2020'][h])
    dataframe['Consumo 05/2020'][h] = float(dataframe['Consumo 05/2020'][h])
    dataframe['Consumo 06/2020'][h] = float(dataframe['Consumo 06/2020'][h])
    dataframe['Consumo 07/2020'][h] = float(dataframe['Consumo 07/2020'][h])
    dataframe['Consumo 08/2020'][h] = float(dataframe['Consumo 08/2020'][h])
    dataframe['Consumo 09/2020'][h] = float(dataframe['Consumo 09/2020'][h])
    dataframe['Consumo 10/2020'][h] = float(dataframe['Consumo 10/2020'][h])
    dataframe['Consumo 11/2020'][h] = float(dataframe['Consumo 11/2020'][h])
    dataframe['Consumo 12/2020'][h] = float(dataframe['Consumo 12/2020'][h])

print(type(dataframe['Consumo 05/2020'][4]))

meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
for i in range(0, len(dataframe)):
    empresa = dataframe['Empresa'][i]
    email = dataframe["Email do responsavel"][i]

    clf = tree.DecisionTreeClassifier()
    # consumo = [janeiro, fevereiro, marco, abril, maio, junho, julho, agosto, setembro, outubro, novembro, dezembro]

features = dataframe.loc[0:102, ['Consumo 01/2020', 'Consumo 02/2020', 'Consumo 03/2020', 'Consumo 04/2020',
                                 'Consumo 05/2020', 'Consumo 06/2020', 'Consumo 07/2020', 'Consumo 08/2020',
                                 'Consumo 09/2020', 'Consumo 10/2020', 'Consumo 11/2020', 'Consumo 12/2020']]

labels = dataframe['Consumo 04/2020']

qtd_linhas = len(dataframe)
qtd_linhas_treino = round(.70 * qtd_linhas)
qtd_linhas_teste = qtd_linhas - qtd_linhas_treino
qtd_linhas_validacao = qtd_linhas - 1

k_best_features = SelectKBest(k='all')
k_best_features.fit_transform(features, labels)
k_best_features_scores = k_best_features.scores_
raw_pairs = zip(meses[1:], k_best_features_scores)
ordered_pairs = list(reversed(sorted(raw_pairs, key=lambda x: x[1])))

k_best_features_final = dict(ordered_pairs[:15])
best_features = k_best_features_final.keys()


features = dataframe.loc[:, ['Consumo 01/2020', 'Consumo 02/2020', 'Consumo 03/2020', 'Consumo 04/2020',
                                 'Consumo 05/2020', 'Consumo 06/2020', 'Consumo 07/2020', 'Consumo 08/2020',
                                 'Consumo 09/2020', 'Consumo 10/2020', 'Consumo 11/2020', 'Consumo 12/2020']]

X_train = features[:qtd_linhas_treino]
X_test = features[qtd_linhas_treino:qtd_linhas_treino + qtd_linhas_teste - 1]

y_train = labels[:qtd_linhas_treino]
y_test = labels[qtd_linhas_treino:qtd_linhas_treino + qtd_linhas_teste -1]

print(len(X_train), len(y_train))
print(len(X_test), len(y_test))

scaler = MinMaxScaler()
X_train_scale = scaler.fit_transform(X_train)
X_test_scale = scaler.transform(X_test)

lr = linear_model.LinearRegression()
lr.fit(X_train_scale, y_train)
pred = lr.predict(X_test_scale)
cd = r2_score(y_test, pred)
print(f'Coeficiente de determinação:{cd * 100:.2f}')

rn = MLPRegressor()

parameter_space = {
    'hidden_layer_sizes': [(i,) for i in list(range(1, 21))],
    'activation': ['tanh', 'relu'],
    'solver': ['sgd', 'adam', 'lbfgs'],
    'alpha': [0.0001, 0.05],
    'learning_rate': ['constant', 'adaptive']
}

search = GridSearchCV(rn, parameter_space, n_jobs=1, cv=5)
search.fit(X_train_scale, y_train)
clf = search.best_estimator_
pred = search.predict(X_test_scale)

cd = search.score(X_test_scale, y_test)
valor_novo = features.tail(1)
previsao = scaler.transform(valor_novo)
predi = lr.predict(previsao)
print(valor_novo)

print(f'Coeficiente de determinação:{cd * 100:.2f}')

# X_train = features[:qtd_linhas_treino]
#https://www.youtube.com/watch?v=rAYFuFNOCic&ab_channel=soumilshah1995
# print(qtd_linhas)


def send_email(destino, anexo):
    outlook = win32.Dispatch('outlook.application')
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

# new.to_csv('new.csv')


# for i in range(0, len(dataframe)):
#     empresa = dataframe['Empresa'][i]
#     email = dataframe["Email do responsavel"][i]
#     janeiro = float(dataframe['Consumo 01/2020'][i].replace(',', '.'))
#     fevereiro = float(dataframe['Consumo 02/2020'][i].replace(',', '.'))
#     marco = float(dataframe['Consumo 03/2020'][i].replace(',', '.'))
#     abril = float(dataframe['Consumo 04/2020'][i].replace(',', '.'))
#     maio = float(dataframe['Consumo 05/2020'][i].replace(',', '.'))
#     junho = float(dataframe['Consumo 06/2020'][i].replace(',', '.'))
#     julho = float(dataframe['Consumo 07/2020'][i].replace(',', '.'))
#     agosto = float(dataframe['Consumo 08/2020'][i].replace(',', '.'))
#     setembro = float(dataframe['Consumo 09/2020'][i].replace(',', '.'))
#     outubro = float(dataframe['Consumo 10/2020'][i].replace(',', '.'))
#     novembro = float(dataframe['Consumo 11/2020'][i].replace(',', '.'))
#     dezembro = float(dataframe['Consumo 12/2020'][i].replace(',', '.'))
#
#     consumo = [janeiro, fevereiro, marco, abril, maio, junho, julho, agosto, setembro, outubro, novembro, dezembro]

# plt.xlabel(meses)
# plt.ylabel(consumo)


# plt.plot(consumo, label='2020', color='red')
# plt.plot(consumo, label='2021', color='green')
# plt.plot(meses, consumo, consumo, color='#427314')
# plt.title(f"{empresa} - Ano 2020")
# plt.plot()
# plt.legend()
#
# plt.show()

# plt.grid()
# fig = plt.savefig(f"{path}/{empresa}.png")
#
# #send_email(email, f"{path}/{empresa}.png")
# print("O e-mail foi enviado para: ", email)
# os.remove(f"{path}/{empresa}.png")

# fig, ax = plt.subplots()
# x = np.arange(len(meses))
# width = 0.35
# rects1 = ax.bar(x - width / 2, meses, width, label='Men')
# rects2 = ax.bar(x + width / 2, consumo, width, label='Women')
# plt.bar(meses, consumo)
# plt.title(f"{empresa} - Ano 2020")

# if __name__ == '__main__':
#     print_hi('PyCharm')
# meses_file = file_2[['Consumo 01/2020', 'Consumo 02/2020', 'Consumo 03/2020', 'Consumo 04/2020',
#                      'Consumo 05/2020', 'Consumo 06/2020', 'Consumo 07/2020', 'Consumo 08/2020',
#                      'Consumo 09/2020', 'Consumo 10/2020', 'Consumo 11/2020', 'Consumo 12/2020']]
