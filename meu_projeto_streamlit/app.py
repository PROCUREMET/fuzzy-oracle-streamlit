import pandas as pd
import unicodedata
from fuzzywuzzy import fuzz
import streamlit as st
import oracledb
import io
import os

def normalizar(texto):
    if pd.isna(texto):
        return ""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', str(texto))
        if not unicodedata.combining(c)
    ).upper()

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

st.title("🔍 Sistema de Correspondência de Itens")

# Criação das abas
abas = ["Busca & Cadastro", "Itens Pendentes para Cadastro"]
tab_busca, tab_pendentes = st.tabs(abas)

# --------- ABA 1: Busca & Cadastro ---------
with tab_busca:
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

                buffer = io.BytesIO()
                resultado_df.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)

                st.download_button(
                    label="📥 Baixar resultado Excel",
                    data=buffer,
                    file_name="resultado_fuzzy.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ Insira ao menos um item para buscar.")

    st.subheader("📝 Formulário para cadastrar novo item")

    with st.form("form_cadastro_item"):
        descricao = st.text_input("Descrição do Item *")
        unidade = st.text_input("Unidade de Medida *")
        dimensoes = st.text_input("Dimensões (opcional)")
        marca = st.text_input("Marca (opcional)")

        enviar = st.form_submit_button("📤 Enviar para cadastro")

    if enviar:
        if not descricao.strip() or not unidade.strip():
            st.error("❌ Os campos 'Descrição do Item' e 'Unidade de Medida' são obrigatórios.")
        else:
            caminho_arquivo = "itens_para_cadastro.xlsx"

            if os.path.exists(caminho_arquivo):
                df_existente = pd.read_excel(caminho_arquivo, engine="openpyxl")
            else:
                df_existente = pd.DataFrame(columns=[
                    "Descrição do Item",
                    "Unidade de Medida",
                    "Dimensões",
                    "Marca",
                    "Data de Envio"
                ])

            novo_item = pd.DataFrame({
                "Descrição do Item": [normalizar(descricao.strip())],
                "Unidade de Medida": [normalizar(unidade.strip())],
                "Dimensões": [normalizar(dimensoes.strip())],
                "Marca": [normalizar(marca.strip())],
                "Data de Envio": [pd.Timestamp.now()]
            })

            df_final = pd.concat([df_existente, novo_item], ignore_index=True)
            df_final.to_excel(caminho_arquivo, index=False, engine="openpyxl")

            st.success("✅ Item enviado com sucesso para a planilha de cadastro.")
            st.info(f"📁 Arquivo salvo em: `{os.path.abspath(caminho_arquivo)}`")

            # Mostrar preview da planilha diretamente na aba
            with st.expander("📄 Ver Itens Pendentes Salvos"):
                st.dataframe(df_final)

# --------- ABA 2: Itens Pendentes ---------
with tab_pendentes:
    st.header("📋 Itens Pendentes para Cadastro")
    caminho_arquivo = "itens_para_cadastro.xlsx"

    if os.path.exists(caminho_arquivo):
        try:
            df_pendentes = pd.read_excel(caminho_arquivo, engine="openpyxl")

            # Normaliza as colunas principais para visualização também
            for col in ["Descrição do Item", "Unidade de Medida", "Dimensões", "Marca"]:
                if col in df_pendentes.columns:
                    df_pendentes[col] = df_pendentes[col].apply(normalizar)

            if not df_pendentes.empty:
                st.dataframe(df_pendentes)
            else:
                st.info("Nenhum item pendente para cadastro encontrado.")
        except Exception as e:
            st.error(f"Erro ao carregar a planilha: {e}")
    else:
        st.info("Nenhum arquivo de itens pendentes encontrado.")
