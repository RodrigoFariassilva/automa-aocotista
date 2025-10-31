import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import io

# ============================
# Funções utilitárias
# ============================
def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

def normalize_column(df, col_name):
    if col_name in df.columns:
        df[col_name] = df[col_name].astype(str).str.upper().str.strip()
        df[col_name] = df[col_name].apply(remove_accents)
    return df

def remove_chars_and_terms(df, col_name):
    chars_and_terms_to_remove = [' ', '.', '/', '-', 'SA', 'LTDA']
    if col_name in df.columns:
        for char_or_term in chars_and_terms_to_remove:
            df[col_name] = df[col_name].str.replace(char_or_term, '', case=False, regex=False)
        df[col_name] = df[col_name].apply(remove_accents)
    return df

# ============================
# Funções de transformação
# ============================
def selecionar_colunas(df):
    colunas_necessarias = ['TITULAR', 'DT.TRANSAÇÃO', 'APLICAR', 'RESGATAR', 'FUNDO']
    df = df[colunas_necessarias].copy()
    df['TITULAR'] = df['TITULAR'].astype(str)
    df['DT.TRANSAÇÃO'] = df['DT.TRANSAÇÃO'].astype(str)
    df['APLICAR'] = pd.to_numeric(df['APLICAR'], errors='coerce')
    df['RESGATAR'] = pd.to_numeric(df['RESGATAR'], errors='coerce')
    return df

def titulares(df_transacoes):
    df_transacoes = df_transacoes[~df_transacoes['TITULAR'].str.contains("TOTAL DA MOVIMENTAÇÃO:", na=False)].copy()
    df_transacoes['TITULAR'] = df_transacoes['TITULAR'].str[6:].str.strip()
    return df_transacoes

def left_merge(df_transacoes, cotistas):
    merged_df = df_transacoes.merge(cotistas, left_on='TITULAR', right_on='Nome', how='left')
    merged_df = merged_df.drop(columns='Nome')
    merged_df['Cliente'] = merged_df['Cliente'].astype(str).apply(lambda x: x.split('.')[0])
    return merged_df

def processar_transacoes(df):
    df['Valor'] = df['APLICAR'].combine_first(df['RESGATAR'])
    df['Transacao'] = np.where(df['APLICAR'].isnull(), 'R', 'A')
    df = df.drop(['APLICAR', 'RESGATAR', 'TITULAR'], axis=1)
    return df

def adicionar_id_fundo(df_movimentacoes, fundos_ids):
    df_movimentacoes = normalize_column(df_movimentacoes, 'FUNDO')
    df_movimentacoes['ID'] = df_movimentacoes['FUNDO'].map(fundos_ids).fillna(20711).astype(int)
    df_movimentacoes['ID_Fundo'] = df_movimentacoes['ID'].astype(str)
    return df_movimentacoes

def reordenar_colunas(df):
    nova_ordem = ['ID_Fundo', 'Cliente', 'Transacao', 'DT.TRANSAÇÃO', 'Valor']
    return df.reindex(columns=nova_ordem)

def formatar_data(df):
    df['DT.TRANSAÇÃO'] = df['DT.TRANSAÇÃO'].str.replace('.', '/')
    return df

def ajustar_valor(df):
    df['Valor'] = df['Valor'].fillna(0).apply(lambda val: "{:015.2f}".format(val).replace('.', ','))
    return df

def dataframe_para_prn(df):
    posicoes = [0, 15, 34, 44, 69]
    colunas = ['ID_Fundo', 'Cliente', 'Transacao', 'DT.TRANSAÇÃO', 'Valor']
    output = io.StringIO()
    for _, row in df.iterrows():
        linha = [' '] * 100
        for i, col in enumerate(colunas):
            valor = str(row[col])
            for j, char in enumerate(valor):
                if posicoes[i] + j < len(linha):
                    linha[posicoes[i] + j] = char
        output.write(''.join(linha).rstrip() + '\n')
    return output.getvalue()

# ============================
# Função para leitura de arquivos de transações
# ============================
def ler_arquivo_transacoes(file):
    try:
        content = file.read().decode('utf-16')
        df = pd.read_csv(io.StringIO(content), sep='\t')
    except Exception:
        file.seek(0)
        df = pd.read_excel(file, engine='xlrd')
    return df

# ============================
# Interface Streamlit
# ============================
st.set_page_config(
    page_title="Automação Cotista",
    page_icon="assets/logo.ico",
)

st.title("Processador de Movimentações Financeiras")
cotistas_path = "assets/Lista Cotistas.csv"
cotistas = pd.read_csv(cotistas_path, delimiter=';')
cotistas = remove_chars_and_terms(cotistas, 'Nome')

transacoes_files = st.file_uploader("Upload dos Arquivos de Transações (.xls, .xlsx)", type=["xls", "xlsx"], accept_multiple_files=True)

if transacoes_files:
    cotistas = pd.read_csv(cotistas_file, delimiter=';')
    cotistas = remove_chars_and_terms(cotistas, 'Nome')

    dfs = []
    arquivos = transacoes_files if isinstance(transacoes_files, list) else [transacoes_files]

    for file in arquivos:
        try:
            df = ler_arquivo_transacoes(file)
            df = selecionar_colunas(df)
            dfs.append(df)
        except Exception as e:
            st.error(f"Erro ao ler {file.name}: {e}")

    df_transacoes = pd.concat(dfs, ignore_index=True)
    df_transacoes = titulares(df_transacoes)
    df_transacoes = remove_chars_and_terms(df_transacoes, 'TITULAR')
    cotistas = remove_chars_and_terms(cotistas, 'Nome')
    df_transacoes = left_merge(df_transacoes, cotistas)
    df_movimentacoes = processar_transacoes(df_transacoes)

    fundos_ids_padrao = {
        'FIDCSENIOR': 20711,
        'FIDCMEZ1': 20731,
        'FIDCMEZ2': 20732,
        'FIDCMEZ3': 20733,
        'FIDCMEZ4': 20734,
        'FIDCMEZ5': 20735
    }

    df_movimentacoes = adicionar_id_fundo(df_movimentacoes, fundos_ids_padrao)
    df_movimentacoes = reordenar_colunas(df_movimentacoes)
    df_movimentacoes = formatar_data(df_movimentacoes)
    df_movimentacoes = ajustar_valor(df_movimentacoes)

    st.subheader("Prévia dos dados processados")
    st.dataframe(df_movimentacoes.head())

    prn_content = dataframe_para_prn(df_movimentacoes)
    st.download_button("Baixar Arquivo PRN", prn_content, file_name="movimentacoes.prn", mime="text/plain")
else:
    st.info("Por favor, envie os arquivos necessários para iniciar o processamento.")
