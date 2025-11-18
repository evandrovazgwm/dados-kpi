import streamlit as st
import pandas as pd
import io

def calcular_ocorrencias(df, coluna_analise='B'):
    """
    Calcula as ocorrências na coluna especificada e aplica as regras de subtração.
    """
    if coluna_analise not in df.columns:
        # Tenta inferir a coluna pelo índice caso o nome não seja 'B' (como em alguns arquivos sem cabeçalho)
        try:
            coluna_analise = df.columns[1] 
        except IndexError:
            st.error(f"A coluna '{coluna_analise}' não foi encontrada no arquivo.")
            return None

    # Conta as ocorrências de cada string na coluna
    contagem = df[coluna_analise].value_counts()
    
    resultados = {}

    # --- Regra 1: W_In - W_Off ---
    w_in = contagem.get('W_In', 0)
    w_off = contagem.get('W_Off', 0)
    resultados['W_Net'] = w_in - w_off
    
    # --- Regra 2: Painting_In - Painting_Out ---
    painting_in = contagem.get('Painting_In', 0)
    painting_out = contagem.get('Painting_Out', 0)
    resultados['Painting_Net'] = painting_in - painting_out

    # --- Regra 3: PBS_IN - PBS_Off ---
    pbs_in = contagem.get('PBS_IN', 0)
    pbs_off = contagem.get('PBS_Off', 0)
    resultados['PBS_Net'] = pbs_in - pbs_off
    
    return resultados, contagem

# --- Interface Streamlit ---

st.title("📊 Calculadora de Ocorrências em Colunas Excel")
st.markdown("Arraste e solte o arquivo Excel para calcular a diferença de ocorrências na **Coluna B**.")

# Widget de upload de arquivo
uploaded_file = st.file_uploader(
    "Escolha um arquivo Excel (.xlsx ou .xls)", 
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    
    try:
        # Leitura do arquivo Excel para um DataFrame
        df = pd.read_excel(uploaded_file, header=None) # header=None assume que a primeira linha não é um cabeçalho
        
        st.success("Arquivo carregado com sucesso!")
        
        # O Streamlit nomeia as colunas por índice quando header=None, então a Coluna B é o índice 1 (0-base)
        coluna_analise = df.columns[1] # Seleciona a segunda coluna (Coluna B)

        st.subheader("Primeiras linhas do arquivo (Colunas A, B, C...)")
        # Exibe as primeiras 5 linhas para o usuário confirmar o carregamento
        st.dataframe(df.head())
        
        st.subheader("🔍 Resultados do Cálculo")

        # Chama a função de cálculo
        resultados, contagem_total = calcular_ocorrencias(df, coluna_analise=coluna_analise)
        
        if resultados:
            
            # Formata os resultados para exibição em uma tabela
            resultados_df = pd.DataFrame(
                list(resultados.items()), 
                columns=['Item', 'Diferença (In - Out/Off)']
            )

            # Exibe a tabela de resultados com destaque
            st.dataframe(
                resultados_df, 
                hide_index=True, 
                use_container_width=True
            )
            
            st.markdown("---")
            
            st.subheader("Contagem Detalhada de Todas as Ocorrências")
            # Exibe a contagem completa para referência
            st.dataframe(
                contagem_total.to_frame(name='Total de Ocorrências'), 
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Ocorreu um erro durante o processamento do arquivo: {e}")
