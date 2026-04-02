import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Automação de Fechamento", layout="wide")

st.title("📊 Automação de Conferência ICMS/IPI")
st.markdown("Arraste sua planilha aqui para processar as colunas de conferência automaticamente.")

# Função auxiliar para inserir coluna ao lado de outra
def inserir_coluna_ao_lado(df, coluna_referencia, nova_coluna, valores):
    if coluna_referencia in df.columns:
        idx = df.columns.get_loc(coluna_referencia) + 1
        df.insert(idx, nova_coluna, valores)
    else:
        df[nova_coluna] = valores # Fallback se não achar a coluna

uploaded_file = st.file_uploader("Upload da Planilha Excel", type="xlsx")

if uploaded_file:
    try:
        # 1. Carregar as abas
        df_detalhe = pd.read_excel(uploaded_file, sheet_name='Detalhamento')
        df_grade = pd.read_excel(uploaded_file, sheet_name='Grade')

        st.info("Planilha carregada com sucesso! Processando regras...")

        # --- PREPARAÇÃO DA ABA GRADE (PROCX por índice de colunas D, K, R) ---
        # Excel Colunas: D=3, K=10, R=17 (No Python começa do 0)
        df_grade_subset = df_grade.iloc[:, [3, 10, 17]].copy()
        df_grade_subset.columns = ['CHAVE_GRADE', 'VAL_ICMS_GRADE', 'VAL_IPI_GRADE']

        # Join (Simulação do PROCX de Chave para Chave)
        df_merged = pd.merge(
            df_detalhe, 
            df_grade_subset, 
            left_on='CHAVE', 
            right_on='CHAVE_GRADE', 
            how='left'
        )

        # Tratamento de valores nulos ou textos para evitar erro em cálculos matemáticos
        df_merged['ALIQUOTA ICMS'] = pd.to_numeric(df_merged['ALIQUOTA ICMS'], errors='coerce').fillna(0)
        df_merged['VAL_ICMS_GRADE'] = pd.to_numeric(df_merged['VAL_ICMS_GRADE'], errors='coerce').fillna(0)
        df_merged['BASE DE CALCULO ICMS'] = pd.to_numeric(df_merged['BASE DE CALCULO ICMS'], errors='coerce').fillna(0)
        df_merged['VALOR ICMS'] = pd.to_numeric(df_merged['VALOR ICMS'], errors='coerce').fillna(0)
        
        df_merged['ALIQUOTA IPI'] = pd.to_numeric(df_merged['ALIQUOTA IPI'], errors='coerce').fillna(0)
        df_merged['VAL_IPI_GRADE'] = pd.to_numeric(df_merged['VAL_IPI_GRADE'], errors='coerce').fillna(0)
        df_merged['BASE DE CALCULO IPI'] = pd.to_numeric(df_merged['BASE DE CALCULO IPI'], errors='coerce').fillna(0)
        df_merged['VALOR IPI'] = pd.to_numeric(df_merged['VALOR IPI'], errors='coerce').fillna(0)

        # --- LÓGICA DO ICMS ---
        aliq_icms_grade = df_merged['VAL_ICMS_GRADE']
        conf_aliq_icms = df_merged['ALIQUOTA ICMS'] == df_merged['VAL_ICMS_GRADE']
        conf_val_icms = (df_merged['BASE DE CALCULO ICMS'] * (df_merged['ALIQUOTA ICMS'] / 100)) - df_merged['VALOR ICMS']

        # Inserindo as colunas no lugar certo (uma ao lado da outra)
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA ICMS', 'ALÍQUOTA ICMS GRADE', aliq_icms_grade)
        inserir_coluna_ao_lado(df_detalhe, 'ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', conf_aliq_icms)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR ICMS', 'CONFERÊNCIA VALOR ICMS', conf_val_icms)

        # --- LÓGICA DO IPI ---
        aliq_ipi_grade = df_merged['VAL_IPI_GRADE']
        conf_aliq_ipi = df_merged['ALIQUOTA IPI'] == df_merged['ALIQUOTA IPI GRADE']
        conf_val_ipi = (df_merged['BASE DE CALCULO IPI'] * (df_merged['ALIQUOTA IPI'] / 100)) - df_merged['VALOR IPI']

        # Inserindo as colunas no lugar certo
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI', 'ALIQUOTA IPI GRADE', aliq_ipi_grade)
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', conf_aliq_ipi)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR IPI', 'CONFERÊNCIA VALOR IPI', conf_val_ipi)

        st.success("✅ Processamento concluído com sucesso!")
        st.dataframe(df_detalhe.head())

        # --- DOWNLOAD DA PLANILHA ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_detalhe.to_excel(writer, index=False, sheet_name='Detalhamento')
            df_grade.to_excel(writer, index=False, sheet_name='Grade')
        
        st.download_button(
            label="📥 Baixar Planilha Processada",
            data=output.getvalue(),
            file_name="Fechamento_Conferido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
