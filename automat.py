import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Automação de Fechamento", layout="wide")

st.title("📊 Automação de Conferência ICMS/IPI")
st.markdown("Arraste sua planilha aqui para processar as colunas de conferência automaticamente.")

# Função para inserir coluna em posição específica
def inserir_coluna_ao_lado(df, coluna_referencia, nova_coluna, valores):
    if coluna_referencia in df.columns:
        idx = df.columns.get_loc(coluna_referencia) + 1
        df.insert(idx, nova_coluna, valores)
    else:
        df[nova_coluna] = valores

uploaded_file = st.file_uploader("Upload da Planilha Excel", type="xlsx")

if uploaded_file:
    try:
        # 1. Carregar as abas (Ignorando espaços extras nos nomes das abas)
        xls = pd.ExcelFile(uploaded_file)
        aba_detalhe = [s for s in xls.sheet_names if 'detalhamento' in s.lower()][0]
        aba_grade = [s for s in xls.sheet_names if 'grade' in s.lower()][0]
        
        df_detalhe = pd.read_excel(uploaded_file, sheet_name=aba_detalhe)
        df_grade = pd.read_excel(uploaded_file, sheet_name=aba_grade)

        st.info("Planilha carregada! Processando...")

        # --- PREPARAÇÃO DA ABA GRADE (PROCX) ---
        # Pegando Coluna D (Chave), K (Icms), R (Ipi) pelo índice (0, 1, 2...)
        # D=3, K=10, R=16 (ou 17 dependendo se há colunas vazias, vamos garantir)
        df_grade_resumo = df_grade.iloc[:, [3, 10, 16]].copy() 
        df_grade_resumo.columns = ['CHAVE_BUSCA', 'ICMS_GRADE_VAL', 'IPI_GRADE_VAL']

        # Join (Merge) para trazer os valores da Grade para o Detalhamento
        df_temp = pd.merge(
            df_detalhe[['CHAVE', 'ALIQUOTA ICMS', 'BASE DE CALCULO ICMS', 'VALOR ICMS', 'ALIQUOTA IPI', 'BASE DE CALCULO IPI', 'VALOR IPI']], 
            df_grade_resumo, 
            left_on='CHAVE', 
            right_on='CHAVE_BUSCA', 
            how='left'
        )

        # Converter colunas para números (evita erro de cálculo se houver texto ou vazio)
        cols_para_converter = [
            'ALIQUOTA ICMS', 'ICMS_GRADE_VAL', 'BASE DE CALCULO ICMS', 'VALOR ICMS',
            'ALIQUOTA IPI', 'IPI_GRADE_VAL', 'BASE DE CALCULO IPI', 'VALOR IPI'
        ]
        for col in cols_para_converter:
            df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0)

        # --- CÁLCULOS ICMS ---
        aliq_icms_grade_valores = df_temp['ICMS_GRADE_VAL']
        conf_aliq_icms_valores = df_temp['ALIQUOTA ICMS'] == df_temp['ICMS_GRADE_VAL']
        conf_val_icms_valores = (df_temp['BASE DE CALCULO ICMS'] * (df_temp['ALIQUOTA ICMS'] / 100)) - df_temp['VALOR ICMS']

        # --- CÁLCULOS IPI ---
        aliq_ipi_grade_valores = df_temp['IPI_GRADE_VAL']
        conf_aliq_ipi_valores = df_temp['ALIQUOTA IPI'] == df_temp['IPI_GRADE_VAL']
        conf_val_ipi_valores = (df_temp['BASE DE CALCULO IPI'] * (df_temp['ALIQUOTA IPI'] / 100)) - df_temp['VALOR IPI']

        # --- INSERÇÃO DAS COLUNAS NO DF ORIGINAL ---
        # ICMS
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA ICMS', 'ALÍQUOTA ICMS GRADE', aliq_icms_grade_valores)
        inserir_coluna_ao_lado(df_detalhe, 'ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', conf_aliq_icms_valores)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR ICMS', 'CONFERÊNCIA VALOR ICMS', conf_val_icms_valores)

        # IPI
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI', 'ALIQUOTA IPI GRADE', aliq_ipi_grade_valores)
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', conf_aliq_ipi_valores)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR IPI', 'CONFERÊNCIA VALOR IPI', conf_val_ipi_valores)

        st.success("✅ Tudo pronto!")
        st.dataframe(df_detalhe.head(10))

        # --- BOTÃO DE DOWNLOAD ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_detalhe.to_excel(writer, index=False, sheet_name='Detalhamento')
            df_grade.to_excel(writer, index=False, sheet_name='Grade')
        
        st.download_button(
            label="📥 Baixar Planilha Processada",
            data=output.getvalue(),
            file_name="Resultado_Conferencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
        st.info("Dica: Verifique se as abas se chamam 'Detalhamento' e 'Grade' e se a coluna 'CHAVE' existe nas duas.")
