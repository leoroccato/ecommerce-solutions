import csv
import re
import pandas as pd
import openpyxl
import xlrd
import numpy as np

csv_path = 'products_export.csv'
excel_path = 'visualizacao.xlsx'
excelf_path = 'fornecedor.xls'
excelf_path1 = 'fornecedor1.xls'
excelf_path2 = 'fornecedor2.xls'

valores = []
nomes = []


def update_title(row):
    title = row['Title']
    parts = [row['Option1 Value'], row['Option2 Value'], row['Option3 Value']]
    # Filtrar partes que não são None e não são NaN
    parts = [part for part in parts if pd.notna(part)]
    if parts:
        title += ' - ' + ' / '.join(parts)
    return title

def merge_planilhas(df1, df2):
    result = pd.merge(df1, df2[['Product Name', 'Total cost']], left_on='Title', right_on='Product Name', how='left')
    result = result[['Title', 'Variant Price', 'Total cost']]
    result.to_excel(excel_path, index=False, engine='openpyxl')
    print(result)

def planilha_fornecedor(excelf_path, excelf_path1, excelf_path2):
    df20 = pd.read_excel(excelf_path)
    df21 = pd.read_excel(excelf_path1)
    df22 = pd.read_excel(excelf_path2)
    df2 = pd.concat([df20, df21, df22], ignore_index=True)
    df2 = df2[['Product Name', 'Quantity', 'Total cost']]
    df2 = df2[df2['Quantity'] == 1]
    return df2


def remover_duplicatas(group):
    if group['Variant Price'].nunique() == 1:
        return group.iloc[:1]
    else:
        return group


def tratamento_csv(csv_path, excel_path):

    df1 = pd.read_csv(csv_path)
    df1 = df1[['Title', 'Option1 Name', 'Option1 Value', 'Option2 Name', 'Option2 Value', 'Option3 Name', 'Option3 Value', 'Variant Price']]
    df1['Title'] = df1['Title'].fillna(method='ffill')
    df1['Option1 Name'] = df1['Option1 Name'].replace('Title', np.nan)
    df1['Option1 Value'] = df1['Option1 Value'].replace('Default Title', np.nan)

    df1['Title'] = df1.apply(update_title, axis=1)

    df1 = df1.dropna(subset=['Variant Price'])
    #df1 = df1.groupby('Title').apply(remover_duplicatas).reset_index(drop=True)
    #df1.to_excel(excel_path, index=False, engine='openpyxl')
    return df1
    print(df1)


df1 = tratamento_csv(csv_path, excel_path)
df2 = planilha_fornecedor(excelf_path, excelf_path1, excelf_path2)
merge_planilhas(df1, df2)