import pandas as pd
import openpyxl
import unicodedata
import numpy as np
import os
import json

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
# Carregamento de arquivos
# ============================

def carregar_csv(caminho_arquivo):
    try:
        df = pd.read_csv(caminho_arquivo, delimiter=';')
        return df
    except Exception as e:
        print(f"Erro ao carregar CSV: {e}")
        return pd.DataFrame()

def carregar_excel(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo, engine="openpyxl")
        return df
    except Exception as e:
        print(f"Erro ao carregar Excel: {e}")
        return pd.DataFrame()

def process_excel_files(dir_path):
    dfs = {}
    try:
        arquivos = os.listdir(dir_path)
        print("Arquivos encontrados:", arquivos)  # Opcional para debug
        for i, filename in enumerate(arquivos, start=1):
            if filename.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(dir_path, filename)
                try:
                    # Usa openpyxl para .xlsx e xlrd para .xls
                    engine = 'openpyxl' if filename.endswith('.xlsx') else 'xlrd'
                    df = pd.read_excel(file_path, engine=engine)
                    dfs[f'df{i}'] = df
                except Exception as e:
                    print(f"Erro ao ler {file_path}: {e}")
        return dfs
    except Exception as e:
        print(f"Erro ao processar arquivos Excel: {e}")
        return {}

def carregar_fundos_multiplos(diretorio):
    fundos_ids_padrao = {
        'COTADEF1': 20711,
        'COTADEF2': 20731,
        'COTADEF3': 20732,
        'COTADEF4': 20733,
        'COTADEF5': 20734,
        'COTADEF6': 20735
    }

    arquivos_excel = [os.path.join(diretorio, f) for f in os.listdir(diretorio) if f.endswith(('.xls', '.xlsx'))]
    dfs = []
    for arquivo in arquivos_excel:
        try:
            df = pd.read_excel(arquivo, engine='openpyxl' if arquivo.endswith('.xlsx') else None)
            dfs.append(df)
        except Exception as e:
            print(f"Erro ao ler {arquivo}: {e}")
    df_fundos = pd.concat(dfs, ignore_index=True)

    df_fundos.columns = [col.strip().upper() for col in df_fundos.columns]
    df_fundos = normalize_column(df_fundos, 'FUNDO')

    if 'ID_FUNDO' not in df_fundos.columns:
        df_fundos['ID_FUNDO'] = df_fundos['FUNDO'].map(fundos_ids_padrao).fillna(20711).astype(int)

    fundos_ids = dict(zip(df_fundos['FUNDO'], df_fundos['ID_FUNDO']))
    return df_fundos, fundos_ids

# ============================
# Funções de transformação
# ============================

def selecionar_colunas(df):
    df = df[['TITULAR', 'DT.TRANSAÇÃO', 'APLICAR', 'RESGATAR', 'FUNDO']].copy()
    df['TITULAR'] = df['TITULAR'].astype(str)
    df['DT.TRANSAÇÃO'] = df['DT.TRANSAÇÃO'].astype(str)
    df['APLICAR'] = df['APLICAR'].astype(float)
    df['RESGATAR'] = df['RESGATAR'].astype(float)
    return df

def titulares(df_transacoes):
    df_transacoes = df_transacoes[~df_transacoes['TITULAR'].str.contains("TOTAL DA MOVIMENTAÇÃO:")].copy()
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

# ============================
# Formatação final
# ============================

def reordenar_colunas(df):
    nova_ordem = ['ID_Fundo', 'Cliente', 'Transacao', 'DT.TRANSAÇÃO', 'Valor']
    return df.reindex(columns=nova_ordem)

def formatar_data(df):
    df['DT.TRANSAÇÃO'] = df['DT.TRANSAÇÃO'].str.replace('.', '/')
    return df

def ajustar_valor(df):
    df['Valor'] = df['Valor'].apply(lambda val: "{:015.2f}".format(val).replace('.', ','))
    return df

def dataframe_para_prn(df, nome_arquivo):
    # Define as posições iniciais de cada coluna
    posicoes = [0, 15, 34, 44, 69]
    colunas = ['ID_Fundo', 'Cliente', 'Transacao', 'DT.TRANSAÇÃO', 'Valor']

    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        for _, row in df.iterrows():
            linha = [' '] * 100  # linha com 100 espaços (ajustável)
            for i, col in enumerate(colunas):
                valor = str(row[col])
                for j, char in enumerate(valor):
                    if posicoes[i] + j < len(linha):
                        linha[posicoes[i] + j] = char
            f.write(''.join(linha).rstrip() + '\n')


# ============================
# Fluxo principal
# ============================

if __name__ == "__main__":
    caminho_cotistas = r'.\'
    dir_fundos = r'.\'
    dir_path = r'.\'

    if not os.path.exists(caminho_cotistas):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_cotistas}")
    if not os.path.exists(dir_fundos):
        raise FileNotFoundError(f"Diretório não encontrado: {dir_fundos}")
    if not os.path.exists(dir_path):
        raise FileNotFoundError(f"Diretório não encontrado: {dir_path}")

    cotistas = carregar_csv(caminho_cotistas)
    cotistas = remove_chars_and_terms(cotistas, 'Nome')

    dfs = process_excel_files(dir_path)
    df_transacoes = pd.concat(dfs.values(), ignore_index=True)
    df_transacoes = selecionar_colunas(df_transacoes)
    df_transacoes = titulares(df_transacoes)
    df_transacoes = remove_chars_and_terms(df_transacoes, 'TITULAR')
    cotistas = remove_chars_and_terms(cotistas, 'Nome')
    df_transacoes = left_merge(df_transacoes, cotistas)

    df_movimentacoes = processar_transacoes(df_transacoes)

    df_fundos, fundos_ids = carregar_fundos_multiplos(dir_fundos)
    df_movimentacoes = adicionar_id_fundo(df_movimentacoes, fundos_ids)
    df_movimentacoes = reordenar_colunas(df_movimentacoes)
    df_movimentacoes = formatar_data(df_movimentacoes)
    df_movimentacoes = ajustar_valor(df_movimentacoes)

    dataframe_para_prn(df_movimentacoes, 'movimentacoes.prn')
