
import pandas as pd
import select
from certifi import where
import oracledb

# Configuração da conexão

oracledb.init_oracle_client(lib_dir=r"C:\oracle\instant_client_21_14")

dsn_tns = oracledb.makedsn(host='dbconnect.megaerp.online', port='4221', service_name='xepdb1')


connection = oracledb.connect(user='TECVERDE', password='Mt2GAcp7KH', dsn=dsn_tns)

# Criação de um cursor

cursor = connection.cursor()


# Executando uma consulta SQL

query = 'select PRO_IN_CODIGO , ALM_IN_CODIGO , LOC_IN_CODIGO , MVS_RE_QUANTIDADE from TECVERDE.EST_MOVSUMARIZADO @Tecverde WHERE ALM_IN_CODIGO IN (16 , 20) and LOC_IN_CODIGO in (1 , 3) and FIL_IN_CODIGO in 3';
cursor.execute(query)


# juntar resultados

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)

# Exportar para um arquivo Excel
df.to_excel('estoque.xlsx', index=False, engine='openpyxl')
#print(df)

cursor.close()
connection.close()



