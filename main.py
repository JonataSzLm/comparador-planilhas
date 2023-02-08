import re
import string

import pandas as pd
from unidecode import unidecode
import nltk

nltk.download("punkt")
from nltk.tokenize import word_tokenize


def clear_string(input_string):
    cleaned_string = re.sub('[%s]' % re.escape(string.punctuation), '', input_string)
    #cleaned_string = re.sub('[%s]' % re.escape(string.digits), '', cleaned_string)
    cleaned_string = re.sub('[^\w\s]', '', cleaned_string)
    cleaned_string = unidecode(cleaned_string)
    cleaned_string = cleaned_string.lower()

    return cleaned_string


if __name__ == '__main__':
    df_nomes_ok = pd.read_excel('./Planilhas/nomes-ok.xlsx', engine='openpyxl')
    df_nomes_cod = pd.read_excel('./Planilhas/nomes-cod.xlsx', engine='openpyxl')

    nomes_ok = df_nomes_ok['Nome'].to_list()
    nomes_cod = df_nomes_cod['Nomes'].to_list()

    result_list = []

    for nome_ok in nomes_ok:
        nome_ok_tokens = word_tokenize(nome_ok)
        nome_ok_tokens = list(map(clear_string, nome_ok_tokens))
        max_instersections = 0
        iFind = None
        for i, nome_cod in enumerate(nomes_cod):
            nome_cod_tokens = word_tokenize(nome_cod)
            nome_cod_tokens = list(map(clear_string, nome_cod_tokens))
            instersections = len(set(nome_ok_tokens).intersection(nome_cod_tokens))
            if instersections > max_instersections:
                max_instersections = instersections
                iFind = i

        result_list.append(iFind)

    result = []
    for i, iresult in enumerate(result_list):
        result_dict = dict()
        result_dict['Nome_ok'] = df_nomes_ok.at[i, 'Nome']
        result_dict['Nome_cod'] = df_nomes_cod.at[iresult, 'Nomes'] if not iresult == None else ''
        result_dict['Id'] = df_nomes_ok.at[i, 'ID']
        result_dict['Cod'] = df_nomes_cod.at[iresult, 'Codigo'] if not iresult == None else ''
        
        result.append(result_dict)

    df_result = pd.DataFrame(result)
    df_result.to_excel('./Planilhas/resultado.xlsx', index=False)

        

        