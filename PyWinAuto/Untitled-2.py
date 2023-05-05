import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy
import itertools

# xl = pd.ExcelFile("./Book1.xlsx")
# print(xl.parse('Sheet1'))

df = pd.read_excel('./GPGT.xlsm', sheet_name='Valeurs')

dataframe = df.iloc[:, 4:6]

print(dataframe)

# datas = []
# for i in df.index:
#     if (df["VALEUR.1"][i] is not numpy.nan):
#         datas.append([df["VALEUR.1"][i], df["LIBELLE.1"][i]])
datas = []

for (i, v) in dataframe.iterrows():
    print(v.values)
for namesTuple in dataframe.itertuples(index=False):
    if (namesTuple[1] is not numpy.nan):
        print(namesTuple[0], namesTuple[1])

# datas = []

# for v in dataframe.index:
#     datas.append([df[df.columns[:4]][v]])

print(datas)

k = [1, 3, 4] + [5, 6, 7]

print(k)
