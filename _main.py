""" 
Robô roda query no sql, de acordo com o dia da semana uma especificação de data diferente,
transforma esse resultado em um Excel e depois popula uma planilha no Google Drive com os dados coletados.

            Criado por: Isabelly Cristine Lopes

            Você pode me encontrar em: 

            Linkedin ->  https://www.linkedin.com/in/isabelly-cristine-lopes-8a9b59204/
            Instagram -> @isabellyloppess
"""

import _conn
from openpyxl import load_workbook
import pandas as pd
import gspread
from google.oauth2 import service_account
import _email

try:
    # Conexão
    _conn.parametriza()
    _conn.dataframe()


    # Drive
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    json_file = "credentials.json"


    def login():
        credentials = service_account.Credentials.from_service_account_file(json_file)
        scoped_credentials = credentials.with_scopes(scopes)
        gc = gspread.authorize(scoped_credentials)
        return gc


    wb = load_workbook('receptivo.xlsx')
    wba = wb.worksheets[0]
    df = pd.DataFrame(wba.values, index=None)

    lista = []

    for row in df[0][:1]:
        lista.append(row)
    for row in df[1][:1]:
        lista.append(row)
    for row in df[2][:1]:
        lista.append(row)
    for row in df[3][:1]:
        lista.append(row)
    for row in df[4][:1]:
        lista.append(row)
    for row in df[5][:1]:
        lista.append(row)
    for row in df[6][:1]:
        lista.append(row)
    for row in df[7][:1]:
        lista.append(row)
    for row in df[8][:1]:
        lista.append(row)
    for row in df[9][:1]:
        lista.append(row)
    print(lista)


    gc = login()
    planilha = gc.open('Planinha')
    planilha = planilha.worksheet('Guia')
    planilha.append_row(lista, value_input_option='USER_ENTERED')

    _email.enviar_email(f'''
                <html>
                    <body>
                        <p> "Robô preencheu com sucesso" </p>
                    </body>
                </html> 
        ''')


except Exception as err:
    _email.enviar_email(f'''
                <html>
                    <body>
                        <p> "Robô não preencheu, erro: {err}" </p>
                    </body>
                </html> 
        ''')


