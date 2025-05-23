import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# 1. Carregar o arquivo
file_path = 'R.I.E.xlsm'  # Coloque o caminho do seu arquivo aqui
sheet_name = 'R.I.E'  # Nome da aba

df = pd.read_excel(file_path, sheet_name=sheet_name)

# 2. Corrigir nome da coluna
df.rename(columns={"Descrição do ItemRELAÇÃO DE INSUMOS ERP": "Descrição"}, inplace=True)

# 3. Limpar dados
df = df.dropna(subset=["Descrição"]).reset_index(drop=True)

# 4. Comparar descrições
vectorizer = TfidfVectorizer().fit_transform(df['Descrição'])
cosine_sim = cosine_similarity(vectorizer)

# 5. Encontrar descrições similares
similar_pairs = []
threshold = 0.90  # Ajuste o nível de "semelhança" aqui (0.90 = 90%)

n = cosine_sim.shape[0]
for i in range(n):
    for j in range(i + 1, n):
        if cosine_sim[i, j] > threshold:
            similar_pairs.append((i, j, cosine_sim[i, j]))

# 6. Montar resultado
duplicated_items = pd.DataFrame([{
    "Item 1": df.loc[i, "Descrição"],
    "Item 2": df.loc[j, "Descrição"],
    "Similaridade": sim
} for i, j, sim in similar_pairs])

# 7. Salvar em Excel
output_path = 'itens_duplicados_similares.xlsx'
duplicated_items.to_excel(output_path, index=False)

print(f"Arquivo gerado com sucesso: {output_path}")