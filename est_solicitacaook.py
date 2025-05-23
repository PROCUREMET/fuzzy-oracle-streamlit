
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

query = 'select ORG_IN_CODIGO, FIL_IN_CODIGO , SOL_IN_CODIGO , SOL_DT_EMISSAO , SOL_BO_APROVAAPPROVO , PROJ_IN_REDUZIDO from TECVERDE.est_solicitacao@Tecverde WHERE EXTRACT(YEAR FROM SOL_DT_EMISSAO) in (2024, 2025)';
cursor.execute(query)


# juntar resultados

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)

# Exportar para um arquivo Excel
df.to_excel('est_solicitacao.xlsx', index=False, engine='openpyxl')
#print(df)

cursor.close()
connection.close()