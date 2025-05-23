import pandas as pd
import select
from certifi import where
import oracledb
import re

# Configuração da conexão
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instant_client_21_14")
dsn_tns = oracledb.makedsn(host='dbconnect.megaerp.online', port='4221', service_name='xepdb1')
connection = oracledb.connect(user='TECVERDE', password='Mt2GAcp7KH', dsn=dsn_tns)

# Criação de um cursor
cursor = connection.cursor()

# Executando uma consulta SQL
query = '''
select 
    ORG_IN_CODIGO, SOL_IN_CODIGO, SOI_IN_CODIGO, UNI_ST_UNIDADE, 
    PRO_IN_CODIGO, SOI_DT_INCLUSAO, SOI_DT_NECESSIDADE, 
    SOI_RE_QUANTIDADESOL, SOI_RE_QTBAIXADA, USU_IN_INCLUSAO, 
    SOI_CH_STATUS, SOI_CH_STATUSNEC, SOI_ST_ESPECIFICACAO, 
    SOI_ST_MOTIVOSOLICITACAO 
from TECVERDE.est_itenssoli@Tecverde 
WHERE EXTRACT(YEAR FROM SOI_DT_INCLUSAO) in (2024, 2025)'''
cursor.execute(query)

# Juntar resultados
rows = cursor.fetchall()
colunas = [desc[0] for desc in cursor.description]
df = pd.DataFrame(rows , columns=colunas)

# 🔧 Remover caracteres ilegais para o Excel
def remove_illegal_chars(val):
    if isinstance(val, str):
        return re.sub(r"[\x00-\x1F]+", "", val)
    return val

df = df.applymap(remove_illegal_chars)

# Exportar para um arquivo Excel
df.to_excel('itenssoli.xlsx', index=False, engine='openpyxl')

# Encerrar conexão
cursor.close()
connection.close()
