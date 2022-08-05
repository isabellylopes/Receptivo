import pyodbc
from datetime import datetime
import pandas as pd

dia = datetime.today().weekday()

# Conexão SQL

server = '*'
database = '*'
login = '*'
senha = '*'

cnxn = pyodbc.connect('DRIVER={SQL SERVER};SERVER=' + server + ';DATABASE=' + database + ';UID=' + login
                      + ';PWD=' + senha)

cursor = cnxn.cursor()

# Parametriza qual query rodar a partir do sia da semana que estamos

def parametriza():
    if dia == 0:
        SuaTabela = ("""USE SeuDatabase 

            DECLARE @DATA AS DATE
            SET @DATA = CAST(GETDATE()-3 AS DATE)
            Set @Data = Cast(@Data as Datetime)
            
            Insert into SuaTabela 
            
            Exec SuaProc @Data  """)

        cursor.execute(SuaTabela)
        cursor.commit()
        x = 'select * from SuaTabela'
        cursor.execute(x)

    else:
        SuaTabela = ("""USE dbMis 

            DECLARE @DATA AS DATE
            SET @DATA = CAST(GETDATE()-1 AS DATE)
            Set @Data = Cast(@Data as Datetime)

            Insert into SuaTabela 

            Exec SuaProc @Data """)

        cursor.execute(SuaTabela)
        cursor.commit()
        x = 'select * from SuaTabela'
        cursor.execute(x)


def dataframe():
    cmd = "select * from SuaTabela"
    results = cursor.execute(cmd).fetchall()

    df = pd.read_sql_query(cmd, cnxn)

    # Tranforma o DataFrame em arquivo Excel
    df.to_excel('receptivo.xlsx', index=None, header=None)

# Removendo as linhas da tabela após o uso
y = 'truncate table SuaTabela'
cursor.execute(y)
