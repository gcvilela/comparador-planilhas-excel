import os
import pandas as pd
import unicodedata


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join([c for c in texto if not unicodedata.combining(c)])


# Caminho dos arquivos
diretorio = os.getcwd()  # Diretório atual de execução
arquivo_1 = os.path.join(diretorio, "1.xlsx")
arquivo_2 = os.path.join(diretorio, "2.xlsx")

if not os.path.exists(arquivo_1) or not os.path.exists(arquivo_2):
    raise FileNotFoundError("Arquivos 1.xlsx e/ou 2.xlsx não encontrados.")

# Leitura
df1 = pd.read_excel(arquivo_1)
df2 = pd.read_excel(arquivo_2)

# Padronizar colunas
df1.columns = df1.columns.str.strip().str.upper()
df2.columns = df2.columns.str.strip().str.upper()

# Localizar colunas exatas
col_paciente = "PACIENTE"
col_senha = "SENHA"
col_usuario = "USUÁRIO"
col_num_guia = "NÚMERO GUIA DA OPERADORA"

# Criar chave com as colunas correspondentes
df1["CHAVE"] = df1.apply(lambda row: normalizar_texto(
    str(int(float(row[col_senha]))) if pd.notna(row[col_senha]) and str(
        row[col_senha]).replace('.', '', 1).isdigit() else ""
) + "|" + normalizar_texto(row[col_paciente]), axis=1)

df2["CHAVE"] = df2.apply(lambda row: normalizar_texto(
    row[col_num_guia]) + "|" + normalizar_texto(row[col_usuario]), axis=1)

# Criar CHAVE_SENHA e CHAVE_GUIA como strings normalizadas
df1["CHAVE_SENHA"] = df1[col_senha].apply(
    lambda x: normalizar_texto(str(int(float(x)))) if pd.notna(
        x) and str(x).replace('.', '', 1).isdigit() else ""
)
df2["CHAVE_GUIA"] = df2[col_num_guia].apply(
    lambda x: normalizar_texto(str(int(float(x)))) if pd.notna(
        x) and str(x).replace('.', '', 1).isdigit() else ""
)

# Garantir que os tipos de dados sejam consistentes
df1["CHAVE_SENHA"] = df1["CHAVE_SENHA"].astype(str)
df2["CHAVE_GUIA"] = df2["CHAVE_GUIA"].astype(str)

# Comparação
chaves_df1 = set(df1["CHAVE"])
chaves_df2 = set(df2["CHAVE"])

comuns = chaves_df1 & chaves_df2
so_df1 = df1[~df1["CHAVE"].isin(chaves_df2)]
so_df2 = df2[~df2["CHAVE"].isin(chaves_df1)]

# Identificar as coincidências entre SENHA e NÚMERO GUIA DA OPERADORA
chaves_senha_guia = set(df1["CHAVE_SENHA"]) & set(df2["CHAVE_GUIA"])

# Criar DataFrame com as coincidências detalhadas
coincidencias_senha_guia = pd.merge(
    df1[df1["CHAVE_SENHA"].isin(chaves_senha_guia)][[
        col_senha, col_paciente, "CHAVE_SENHA"]],
    df2[df2["CHAVE_GUIA"].isin(chaves_senha_guia)][[
        col_num_guia, col_usuario, "CHAVE_GUIA"]],
    left_on="CHAVE_SENHA",
    right_on="CHAVE_GUIA",
    how="inner"
)

# Renomear colunas para clareza
coincidencias_senha_guia.rename(
    columns={
        col_senha: "SENHA (Planilha 1)",
        col_paciente: "PACIENTE (Planilha 1)",
        col_num_guia: "NÚMERO GUIA (Planilha 2)",
        col_usuario: "USUÁRIO (Planilha 2)"
    },
    inplace=True
)

# Coincidências lado a lado
coincidencias_lado_a_lado = pd.merge(
    df1[df1["CHAVE"].isin(comuns)][[col_senha, col_paciente, "CHAVE"]],
    df2[df2["CHAVE"].isin(comuns)][[col_num_guia, col_usuario, "CHAVE"]],
    on="CHAVE",  # Usar "on" porque a coluna é comum
    how="inner"
)

# Obter as chaves encontradas em Coincidências Senha-Guia
chaves_encontradas = set(coincidencias_senha_guia["CHAVE_SENHA"])

# Remover as chaves encontradas de Não Coincidências 1 e 2
so_df1_filtrado = so_df1[~so_df1["CHAVE_SENHA"].isin(chaves_encontradas)]
so_df2_filtrado = so_df2[~so_df2["CHAVE_GUIA"].isin(chaves_encontradas)]

# Exportar para um único arquivo Excel com múltiplas abas
with pd.ExcelWriter("resultado_comparacao.xlsx", engine="openpyxl") as writer:
    coincidencias_lado_a_lado.to_excel(
        writer, sheet_name="Coincidências Lado a Lado", index=False)
    so_df1_filtrado[[col_senha, col_paciente]].reset_index(drop=True).to_excel(
        writer, sheet_name="Não Coincidências 1", index=False)
    so_df2_filtrado[[col_num_guia, col_usuario]].reset_index(drop=True).to_excel(
        writer, sheet_name="Não Coincidências 2", index=False)
    coincidencias_senha_guia.reset_index(drop=True).to_excel(
        writer, sheet_name="Coincidências Senha-Guia", index=False)

print("✅ Comparação realizada com sucesso!")
print("Arquivo gerado: resultado_comparacao.xlsx")
