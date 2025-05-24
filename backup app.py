import pandas as pd
import unicodedata
from fuzzywuzzy import fuzz
import streamlit as st
import oracledb
import io
import os

# Função para normalizar texto
def normalizar(texto):
    if pd.isna(texto):
        return ""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', str(texto))
        if not unicodedata.combining(c)
    ).upper()

# Carrega os dados do banco Oracle com cache
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

# Título
st.title("🔍 Sistema de Correspondência de Itens")

# Controle de abas
if "aba_atual" not in st.session_state:
    st.session_state.aba_atual = "Busca & Cadastro"

abas = ["Busca & Cadastro", "Itens Pendentes para Cadastro"]
tab_busca, tab_pendentes = st.tabs(abas)

# Aba: Busca & Cadastro
if st.session_state.aba_atual == "Busca & Cadastro":
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

        # Formulário de cadastro
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
                    "Descrição do Item": [descricao.strip()],
                    "Unidade de Medida": [unidade.strip()],
                    "Dimensões": [dimensoes.strip()],
                    "Marca": [marca.strip()],
                    "Data de Envio": [pd.Timestamp.now()]
                })

                df_final = pd.concat([df_existente, novo_item], ignore_index=True)
                df_final.to_excel(caminho_arquivo, index=False, engine="openpyxl")

                st.success("✅ Item enviado com sucesso para a planilha de cadastro.")
                st.info(f"📁 Arquivo salvo em: `{os.path.abspath(caminho_arquivo)}`")

        # Botão opcional para ver os itens pendentes
        if st.button("📄 Ver Itens Pendentes"):
            st.session_state.aba_atual = "Itens Pendentes para Cadastro"
            st.rerun()

# Aba: Itens Pendentes
elif st.session_state.aba_atual == "Itens Pendentes para Cadastro":
    with tab_pendentes:
        st.header("📋 Itens Pendentes para Cadastro")
        caminho_arquivo = "itens_para_cadastro.xlsx"

        if os.path.exists(caminho_arquivo):
            df_pendentes = pd.read_excel(caminho_arquivo, engine="openpyxl")
            if not df_pendentes.empty:
                st.dataframe(df_pendentes)
            else:
                st.info("Nenhum item pendente para cadastro encontrado.")
        else:
            st.info("Nenhum arquivo de itens pendentes encontrado.")
