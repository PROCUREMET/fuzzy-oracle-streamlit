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

query = 'select ITP_DT_ENTREGA , sol_in_codigo , soi_in_codigo , pdc_in_codigo , itp_in_sequencia from TECVERDE.est_solicpedido@Tecverde'
cursor.execute(query)


# Fetching the results

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)

# Exportar para um arquivo Excel
df.to_excel('estsolipedido.xlsx', index=False, engine='openpyxl')

cursor.close()
connection.close()

