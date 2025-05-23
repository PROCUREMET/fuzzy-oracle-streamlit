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

query = 'SELECT LAN_IN_CODIGO , LAN_SEQ_IN_CODIGO , CUS_IN_REDUZIDO , CUS_ST_DESCRICAO , LCC_DT_INCLUSAO from TECVERDE.CON_VW_LANCAMENTO_CC@TECVERDE WHERE EXTRACT(YEAR FROM LCC_DT_INCLUSAO) = 2024';
cursor.execute(query)


# Fetching the results

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)

# Exportar para um arquivo Excel
df.to_excel('teste filter123.xlsx', index=False, engine='openpyxl')


cursor.close()
connection.close()


