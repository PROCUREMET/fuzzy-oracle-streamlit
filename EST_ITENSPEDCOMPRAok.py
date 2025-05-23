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

query = 'select PDC_IN_CODIGO , PRO_IN_CODIGO , ITP_RE_QUANTIDADE , ITP_RE_VLUNITARIO , ITP_RE_VLMERCADORIA , ITP_RE_VLTOTAL , itp_st_situacao, ITP_DT_ORCAMENTO, itp_re_vldesc from TECVERDE.est_itenspedcompra@Tecverde WHERE EXTRACT(YEAR FROM ITP_DT_ORCAMENTO) in (2024, 2025)';
cursor.execute(query)


# Fetching the results

rows = cursor.fetchall()

colunas = [desc[0] for desc in cursor.description]


# Criar um DataFrame do pandas com os dados
df = pd.DataFrame(rows , columns=colunas)

# Exportar para um arquivo Excel
df.to_excel('EST_ITENSPEDCOMPRA.xlsx', index=False, engine='openpyxl')
#print(df)
cursor.close()
connection.close()

