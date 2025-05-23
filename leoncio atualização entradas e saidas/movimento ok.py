
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

query = 'select ACAO_IN_CODIGO, LOC_IN_CODIGO , ALM_IN_CODIGO , MVT_ST_NUMDOC, MVT_DT_MOVIMENTO, PRO_IN_CODIGO, MVT_RE_QUANTIDADE, MVT_IN_LANCAM from TECVERDE.EST_MOVIMENTO @Tecverde WHERE ACAO_IN_CODIGO in (844 , 575) and ALM_IN_CODIGO in (1 , 16) AND LOC_IN_CODIGO in (1 , 3) AND EXTRACT(YEAR FROM MVT_DT_MOVIMENTO) in (2024, 2025)'
cursor.execute(query)


# juntar resultados

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)

# Exportar para um arquivo Excel
df.to_excel('movimento atualizado ok.xlsx', index=False, engine='openpyxl')
#print(df)

cursor.close()
connection.close()
