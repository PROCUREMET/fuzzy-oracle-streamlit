
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

query = 'select  MOV_ST_DOCUMENTO, MOV_DT_DATADOCTO, MOV_DT_PRORROGADO, MOV_RE_VALOR, TPD_ST_CODIGO, agn_in_codigo from TECVERDE.fin_movimento@Tecverde WHERE EXTRACT(YEAR FROM MOV_DT_DATADOCTO) in (2024, 2025) AND ACAO_IN_CODIGO in (172, 842, 843, 862)'
cursor.execute(query)


# juntar resultados

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)

# Exportar para um arquivo Excel
df.to_excel('movimentoFin.xlsx', index=False, engine='openpyxl')
#print(df)

cursor.close()
connection.close()
