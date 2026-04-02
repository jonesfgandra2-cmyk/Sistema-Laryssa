import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Automação de Fechamento", layout="wide")

st.title("📊 Automação de Conferência ICMS/IPI")
st.markdown("Arraste sua planilha aqui para processar as colunas de conferência automaticamente.")

def inserir_coluna_ao_lado(df, coluna_referencia, nova_coluna, valores):
    if coluna_referencia in df.columns:
        idx = df.columns.get_loc(coluna_referencia) + 1
        df.insert(idx, nova_coluna, valores)
    else:
        df[nova_coluna] = valores

uploaded_file = st.file_uploader("Upload da Planilha Excel", type="xlsx")

if uploaded_file:
    try:
        # 1. Carregar as abas
        xls = pd.ExcelFile(uploaded_file)
        aba_detalhe = [s for s in xls.sheet_names if 'detalhamento' in s.lower()][0]
        aba_grade = [s for s in xls.sheet_names if 'grade' in s.lower()][0]
        
        # LER TUDO COMO STRING INICIALMENTE PARA EVITAR O ERRO DE "INT TOO LARGE"
        df_detalhe = pd.read_excel(uploaded_file, sheet_name=aba_detalhe, dtype=str)
        df_grade = pd.read_excel(uploaded_file, sheet_name=aba_grade, dtype=str)

        st.info("Planilha carregada! Processando...")

        # --- PREPARAÇÃO DOS DADOS ---
        # Converter apenas as colunas de cálculo para numérico, o resto fica como texto (evita o erro)
        cols_calculo = [
            'ALIQUOTA ICMS', 'BASE DE CALCULO ICMS', 'VALOR ICMS',
            'ALIQUOTA IPI', 'BASE DE CALCULO IPI', 'VALOR IPI'
        ]
        
        for col in cols_calculo:
            if col in df_detalhe.columns:
                df_detalhe[col] = pd.to_numeric(df_detalhe[col], errors='coerce').fillna(0)

        # Preparar a aba Grade (Coluna D=3, K=10, R=16 ou 17)
        # Vamos pegar os valores e converter para número
        df_grade_resumo = df_grade.iloc[:, [3, 10, 16]].copy()
        df_grade_resumo.columns = ['CHAVE_BUSCA', 'ICMS_GRADE_VAL', 'IPI_GRADE_VAL']
        
        df_grade_resumo['ICMS_GRADE_VAL'] = pd.to_numeric(df_grade_resumo['ICMS_GRADE_VAL'], errors='coerce').fillna(0)
        df_grade_resumo['IPI_GRADE_VAL'] = pd.to_numeric(df_grade_resumo['IPI_GRADE_VAL'], errors='coerce').fillna(0)

        # --- SIMULAÇÃO DO PROCX ---
        df_temp = pd.merge(
            df_detalhe[['CHAVE']], 
            df_grade_resumo, 
            left_on='CHAVE', 
            right_on='CHAVE_BUSCA', 
            how='left'
        )

        # --- CÁLCULOS ---
        # ICMS
        aliq_icms_grade = df_temp['ICMS_GRADE_VAL']
        conf_aliq_icms = df_detalhe['ALIQUOTA ICMS'] == df_temp['ICMS_GRADE_VAL']
        conf_val_icms = (df_detalhe['BASE DE CALCULO ICMS'] * (df_detalhe['ALIQUOTA ICMS'] / 100)) - df_detalhe['VALOR ICMS']

        # IPI
        aliq_ipi_grade = df_temp['IPI_GRADE_VAL']
        conf_aliq_ipi = df_detalhe['ALIQUOTA IPI'] == df_temp['IPI_GRADE_VAL']
        conf_val_ipi = (df_detalhe['BASE DE CALCULO IPI'] * (df_detalhe['ALIQUOTA IPI'] / 100)) - df_detalhe['VALOR IPI']

        # --- INSERIR COLUNAS NO DF ORIGINAL ---
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA ICMS', 'ALÍQUOTA ICMS GRADE', aliq_icms_grade)
        inserir_coluna_ao_lado(df_detalhe, 'ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', conf_aliq_icms)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR ICMS', 'CONFERÊNCIA VALOR ICMS', conf_val_icms)

        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI', 'ALIQUOTA IPI GRADE', aliq_ipi_grade)
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', conf_aliq_ipi)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR IPI', 'CONFERÊNCIA VALOR IPI', conf_val_ipi)

        st.success("✅ Tudo pronto!")
        
        # Exibir apenas as primeiras colunas para evitar lentidão no browser
        st.dataframe(df_detalhe.head(50))

        # --- DOWNLOAD ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_detalhe.to_excel(writer, index=False, sheet_name='Detalhamento')
            # Para a aba grade, voltamos ela como estava
            df_grade.to_excel(writer, index=False, sheet_name='Grade')
        
        st.download_button(
            label="📥 Baixar Planilha Processada",
            data=output.getvalue(),
            file_name="Resultado_Conferencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
