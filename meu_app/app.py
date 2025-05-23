import pandas as pd
import unicodedata
from fuzzywuzzy import fuzz
import streamlit as st
import oracledb
import io

# Função para normalizar textos
def normalizar(texto):
    if pd.isna(texto):
        return ""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', str(texto))
        if not unicodedata.combining(c)
    ).upper()

# Função para carregar dados da base Oracle
@st.cache_data(show_spinner=True)
def carregar_dados():
    oracledb.init_oracle_client(lib_dir=r"C:\oracle\instant_client_21_14")
    dsn_tns = oracledb.makedsn(
        host='dbconnect.megaerp.online',
        port='4221',
        service_name='xepdb1'
    )
    connection = oracledb.connect(user='TECVERDE', password='Mt2GAcp7KH', dsn=dsn_tns)
    cursor = connection.cursor()

    query = '''
        SELECT 
            UNI_ST_UNIDADE, 
            PRO_ST_DEFITEM, 
            PRO_IN_CODIGO, 
            PRO_ST_DESCRICAO, 
            GRU_IDE_ST_CODIGO 
        FROM TECVERDE.est_produtos@Tecverde
    '''
    cursor.execute(query)
    rows = cursor.fetchall()
    colunas = [desc[0] for desc in cursor.description]

    df = pd.DataFrame(rows, columns=colunas)
    cursor.close()
    connection.close()

    df["DESCRICAO_NORMALIZADA"] = df["PRO_ST_DESCRICAO"].apply(normalizar)
    return df

# Interface Streamlit
st.title("🔍 Sistema de Correspondência de Itens:")

# Entrada dos itens do usuário
entrada = st.text_area("📋 Cole os itens a serem consultados MEGA (um por linha):")

if st.button("🔎 Buscar na Base de Dados"):
    if entrada.strip():
        with st.spinner("🔄 Conectando ao banco e processando..."):
            df = carregar_dados()
            itens_usuario = [normalizar(i) for i in entrada.splitlines() if i.strip()]
            resultados = []

            for item in itens_usuario:
                melhores = df["DESCRICAO_NORMALIZADA"].apply(lambda x: fuzz.token_sort_ratio(item, x))
                indice_melhor = melhores.idxmax()
                similaridade = melhores.max()
                linha = df.loc[indice_melhor]

                resultados.append({
                    "Item Procurado": item,
                    "Mais Próximo na Base": linha["PRO_ST_DESCRICAO"],
                    "Código": linha["PRO_IN_CODIGO"],
                    "Unidade": linha["UNI_ST_UNIDADE"],
                    "Grupo": linha["GRU_IDE_ST_CODIGO"],
                    "Similaridade": similaridade
                })

            resultado_df = pd.DataFrame(resultados)
            st.success("✅ Busca finalizada!")
            st.dataframe(resultado_df)

            # Exportação para Excel em memória
            buffer = io.BytesIO()
            resultado_df.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)

            # Botão de download
            st.download_button(
                label="📥 Baixar resultado Excel",
                data=buffer,
                file_name="resultado_fuzzy.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Insira ao menos um item para buscar.")
