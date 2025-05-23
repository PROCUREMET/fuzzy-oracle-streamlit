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

query = 'select UNI_ST_UNIDADE , PRO_ST_DEFITEM , pro_in_codigo , pro_st_descricao, gru_ide_st_codigo from TECVERDE.est_produtos@Tecverde'

cursor.execute(query)


# Fetching the results

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)


# Exportar para um arquivo Excel
df.to_csv('est_produto.csv', encoding='utf-8-sig', index=False)
#print(df)

cursor.close()
connection.close()