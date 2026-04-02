import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Automação de Fechamento", layout="wide")

st.title("📊 Automação de Conferência Fiscal Completa")
st.markdown("Processamento de ICMS, IPI, PIS e COFINS com regras de exceção de CFOP para Alíquotas Grade.")

def inserir_coluna_ao_lado(df, coluna_referencia, nova_coluna, valores):
    if coluna_referencia in df.columns:
        idx = df.columns.get_loc(coluna_referencia) + 1
        df.insert(idx, nova_coluna, valores)
    else:
        df[nova_coluna] = valores

uploaded_file = st.file_uploader("Upload da Planilha Excel", type="xlsx")

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        aba_detalhe = [s for s in xls.sheet_names if 'detalhamento' in s.lower()][0]
        aba_grade = [s for s in xls.sheet_names if 'grade' in s.lower()][0]
        
        # Leitura inicial como string para evitar erros de números gigantes
        df_detalhe = pd.read_excel(uploaded_file, sheet_name=aba_detalhe, dtype=str)
        df_grade = pd.read_excel(uploaded_file, sheet_name=aba_grade, dtype=str)

        st.info("Planilha carregada! Aplicando regras de exceção de CFOP...")

        # --- LISTAS DE EXCEÇÕES CFOP (FORÇAR ALÍQUOTA 0 NA GRADE) ---
        cfops_icms_zero = [
            '1301', '1303', '1551', '1556', '1905', '1908', '1911', '1916', '1920', '1933',
            '2551', '2556', '2911', '5152', '5502', '5551', '5906', '5908', '5909', '5911',
            '5921', '6551', '6908', '6911', '6915', '7102'
        ]

        cfops_ipi_zero = [
            '1301', '1303', '1551', '1556', '1905', '1908', '1911', '1916', '1920', '1933',
            '2551', '2556', '2911', '5502', '5551', '5906', '5908', '5909', '5911', '5921',
            '6551', '6908', '6911', '6915', '7102'
        ]

        # --- PREPARAÇÃO DOS DADOS ---
        cols_calculo = [
            'ALIQUOTA ICMS', 'BASE DE CALCULO ICMS', 'VALOR ICMS',
            'ALIQUOTA IPI', 'BASE DE CALCULO IPI', 'VALOR IPI',
            'ALIQUOTA PIS', 'BASE DE CALCULO PIS', 'VALOR PIS',
            'BASE DE CALCULO COFINS', 'ALIQUOTA COFINS', 'VALOR COFINS'
        ]
        for col in cols_calculo:
            if col in df_detalhe.columns:
                df_detalhe[col] = pd.to_numeric(df_detalhe[col], errors='coerce').fillna(0)

        # Preparação Aba Grade (3=Chave, 9=ICMS(J), 16=IPI(Q), 18=PIS(S))
        df_grade_resumo = df_grade.iloc[:, [3, 9, 16, 18]].copy()
        df_grade_resumo.columns = ['CHAVE_BUSCA', 'ICMS_GRADE_VAL', 'IPI_GRADE_VAL', 'PIS_GRADE_VAL']
        
        for col in ['ICMS_GRADE_VAL', 'IPI_GRADE_VAL', 'PIS_GRADE_VAL']:
            df_grade_resumo[col] = pd.to_numeric(df_grade_resumo[col], errors='coerce').fillna(0)

        # Cruzamento
        df_temp = pd.merge(
            df_detalhe[['CHAVE', 'CFOP']], 
            df_grade_resumo, 
            left_on='CHAVE', 
            right_on='CHAVE_BUSCA', 
            how='left'
        )
        
        # Limpeza do CFOP para comparação
        df_temp['CFOP_CLEAN'] = df_temp['CFOP'].astype(str).str.strip()

        # --- APLICANDO EXCEÇÕES ICMS ---
        aliq_icms_grade = df_temp['ICMS_GRADE_VAL'].copy()
        aliq_icms_grade.loc[df_temp['CFOP_CLEAN'].isin(cfops_icms_zero)] = 0

        # --- APLICANDO EXCEÇÕES IPI ---
        aliq_ipi_grade = df_temp['IPI_GRADE_VAL'].copy()
        aliq_ipi_grade.loc[df_temp['CFOP_CLEAN'].isin(cfops_ipi_zero)] = 0

        # --- CÁLCULOS ---
        # ICMS
        conf_aliq_icms = df_detalhe['ALIQUOTA ICMS'] == aliq_icms_grade
        conf_val_icms = (df_detalhe['BASE DE CALCULO ICMS'] * (df_detalhe['ALIQUOTA ICMS'] / 100)) - df_detalhe['VALOR ICMS']

        # IPI
        conf_aliq_ipi = df_detalhe['ALIQUOTA IPI'] == aliq_ipi_grade
        conf_val_ipi = (df_detalhe['BASE DE CALCULO IPI'] * (df_detalhe['ALIQUOTA IPI'] / 100)) - df_detalhe['VALOR IPI']

        # PIS
        aliq_pis_grade = df_temp['PIS_GRADE_VAL']
        conf_aliq_pis = df_detalhe['ALIQUOTA PIS'] == aliq_pis_grade
        conf_val_pis = (df_detalhe['BASE DE CALCULO PIS'] * (df_detalhe['ALIQUOTA PIS'] / 100)) - df_detalhe['VALOR PIS']

        # COFINS
        conf_val_cofins = (df_detalhe['BASE DE CALCULO COFINS'] * (df_detalhe['ALIQUOTA COFINS'] / 100)) - df_detalhe['VALOR COFINS']

        # --- LIMPEZA DE COLUNAS ANTERIORES ---
        cols_novas = [
            'ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', 'CONFERÊNCIA VALOR ICMS', 
            'ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', 'CONFERÊNCIA VALOR IPI',
            'ALIQUOTA PIS GRADE', 'CONFERÊNCIA ALÍQUOTA PIS', 'CONFERÊNCIA VALOR PIS',
            'CONFERÊNCIA VALOR COFINS'
        ]
        for c in cols_novas:
            if c in df_detalhe.columns:
                df_detalhe.drop(columns=[c], inplace=True)

        # --- INSERÇÃO ---
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA ICMS', 'ALÍQUOTA ICMS GRADE', aliq_icms_grade)
        inserir_coluna_ao_lado(df_detalhe, 'ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', conf_aliq_icms)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR ICMS', 'CONFERÊNCIA VALOR ICMS', conf_val_icms)

        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI', 'ALIQUOTA IPI GRADE', aliq_ipi_grade)
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', conf_aliq_ipi)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR IPI', 'CONFERÊNCIA VALOR IPI', conf_val_ipi)

        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA PIS', 'ALIQUOTA PIS GRADE', aliq_pis_grade)
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA PIS GRADE', 'CONFERÊNCIA ALÍQUOTA PIS', conf_aliq_pis)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR PIS', 'CONFERÊNCIA VALOR PIS', conf_val_pis)

        inserir_coluna_ao_lado(df_detalhe, 'VALOR COFINS', 'CONFERÊNCIA VALOR COFINS', conf_val_cofins)

        st.success("✅ Processamento finalizado! Exceções de CFOP aplicadas para ICMS e IPI.")
        st.dataframe(df_detalhe.head(50))

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_detalhe.to_excel(writer, index=False, sheet_name='Detalhamento')
            df_grade.to_excel(writer, index=False, sheet_name='Grade')
        
        st.download_button(label="📥 Baixar Planilha Processada", data=output.getvalue(), file_name="Conferencia_Final_Excecoes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
