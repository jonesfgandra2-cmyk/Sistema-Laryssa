import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Automação de Conferência Fiscal", layout="wide")

st.title("📊 Automação de Conferência Fiscal Completa")
st.markdown("Processamento otimizado para grandes volumes de dados.")

def inserir_coluna_ao_lado(df, coluna_referencia, nova_coluna, valores):
    if coluna_referencia in df.columns:
        idx = df.columns.get_loc(coluna_referencia) + 1
        # Se a coluna já existir, remove para não dar erro de duplicidade
        if nova_coluna in df.columns:
            df.drop(columns=[nova_coluna], inplace=True)
        df.insert(idx, nova_coluna, valores)
    else:
        df[nova_coluna] = valores

uploaded_file = st.file_uploader("Upload da Planilha Excel", type="xlsx")

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        aba_detalhe = [s for s in xls.sheet_names if 'detalhamento' in s.lower()][0]
        aba_grade = [s for s in xls.sheet_names if 'grade' in s.lower()][0]
        
        # Lendo tudo como string para evitar erros de memória e de tipos
        df_detalhe = pd.read_excel(uploaded_file, sheet_name=aba_detalhe, dtype=str)
        df_grade = pd.read_excel(uploaded_file, sheet_name=aba_grade, dtype=str)

        st.info("Planilha carregada! Aplicando regras de negócio...")

        # --- LISTAS DE EXCEÇÕES ---
        cfops_comum = ['1301', '1303', '1551', '1556', '1905', '1908', '1911', '1916', '1920', '1933',
                       '2551', '2556', '2911', '5502', '5551', '5906', '5908', '5909', '5911',
                       '5921', '6551', '6908', '6911', '6915', '7102']
        
        cfops_icms_zero = cfops_comum + ['5152']
        cfops_ipi_zero = cfops_comum
        cfops_pis_zero = cfops_comum + ['5152', '5910', '6910']

        # --- TRATAMENTO DE DADOS (CONVERSÃO NUMÉRICA) ---
        cols_calculo = [
            'ALIQUOTA ICMS', 'BASE DE CALCULO ICMS', 'VALOR ICMS',
            'ALIQUOTA IPI', 'BASE DE CALCULO IPI', 'VALOR IPI',
            'ALIQUOTA PIS', 'BASE DE CALCULO PIS', 'VALOR PIS',
            'BASE DE CALCULO COFINS', 'ALIQUOTA COFINS', 'VALOR COFINS'
        ]
        for col in cols_calculo:
            if col in df_detalhe.columns:
                df_detalhe[col] = pd.to_numeric(df_detalhe[col], errors='coerce').fillna(0)

        # Limpeza do CFOP e Chave
        df_detalhe['CFOP_CLEAN'] = df_detalhe['CFOP'].astype(str).str.strip()
        
        # --- PREPARAÇÃO DA GRADE (EVITANDO DUPLICADOS NO PROCX) ---
        # Índice 3=Chave, 9=ICMS(J), 16=IPI(Q), 18=PIS(S)
        df_grade_resumo = df_grade.iloc[:, [3, 9, 16, 18]].copy()
        df_grade_resumo.columns = ['CHAVE_LOOKUP', 'ICMS_G', 'IPI_G', 'PIS_G']
        
        # IMPORTANTE: Se houver a mesma chave duplicada na Grade, o merge multiplica as linhas.
        # Vamos manter apenas a primeira ocorrência da chave na Grade.
        df_grade_resumo = df_grade_resumo.drop_duplicates(subset=['CHAVE_LOOKUP'])

        for col in ['ICMS_G', 'IPI_G', 'PIS_G']:
            df_grade_resumo[col] = pd.to_numeric(df_grade_resumo[col], errors='coerce').fillna(0)

        # --- CRUZAMENTO (MERGE) ---
        # Unimos as informações da Grade diretamente no df_detalhe
        df_detalhe = pd.merge(df_detalhe, df_grade_resumo, left_on='CHAVE', right_on='CHAVE_LOOKUP', how='left')
        
        # Preencher vazios caso a chave não tenha sido encontrada
        df_detalhe[['ICMS_G', 'IPI_G', 'PIS_G']] = df_detalhe[['ICMS_G', 'IPI_G', 'PIS_G']].fillna(0)

        # --- APLICANDO EXCEÇÕES ---
        # Se CFOP na lista, alíquota Grade vira 0
        df_detalhe.loc[df_detalhe['CFOP_CLEAN'].isin(cfops_icms_zero), 'ICMS_G'] = 0
        df_detalhe.loc[df_detalhe['CFOP_CLEAN'].isin(cfops_ipi_zero), 'IPI_G'] = 0
        df_detalhe.loc[df_detalhe['CFOP_CLEAN'].isin(cfops_pis_zero), 'PIS_G'] = 0

        # --- CÁLCULOS (TUDO NA MESMA TABELA AGORA) ---
        # ICMS
        conf_aliq_icms = df_detalhe['ALIQUOTA ICMS'] == df_detalhe['ICMS_G']
        conf_val_icms = (df_detalhe['BASE DE CALCULO ICMS'] * (df_detalhe['ALIQUOTA ICMS'] / 100)) - df_detalhe['VALOR ICMS']
        
        # IPI
        conf_aliq_ipi = df_detalhe['ALIQUOTA IPI'] == df_detalhe['IPI_G']
        conf_val_ipi = (df_detalhe['BASE DE CALCULO IPI'] * (df_detalhe['ALIQUOTA IPI'] / 100)) - df_detalhe['VALOR IPI']
        
        # PIS
        conf_aliq_pis = df_detalhe['ALIQUOTA PIS'] == df_detalhe['PIS_G']
        conf_val_pis = (df_detalhe['BASE DE CALCULO PIS'] * (df_detalhe['ALIQUOTA PIS'] / 100)) - df_detalhe['VALOR PIS']
        
        # COFINS
        conf_val_cofins = (df_detalhe['BASE DE CALCULO COFINS'] * (df_detalhe['ALIQUOTA COFINS'] / 100)) - df_detalhe['VALOR COFINS']

        # --- INSERÇÃO ---
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA ICMS', 'ALÍQUOTA ICMS GRADE', df_detalhe['ICMS_G'])
        inserir_coluna_ao_lado(df_detalhe, 'ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', conf_aliq_icms)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR ICMS', 'CONFERÊNCIA VALOR ICMS', conf_val_icms)

        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI', 'ALIQUOTA IPI GRADE', df_detalhe['IPI_G'])
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', conf_aliq_ipi)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR IPI', 'CONFERÊNCIA VALOR IPI', conf_val_ipi)

        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA PIS', 'ALIQUOTA PIS GRADE', df_detalhe['PIS_G'])
        inserir_coluna_ao_lado(df_detalhe, 'ALIQUOTA PIS GRADE', 'CONFERÊNCIA ALÍQUOTA PIS', conf_aliq_pis)
        inserir_coluna_ao_lado(df_detalhe, 'VALOR PIS', 'CONFERÊNCIA VALOR PIS', conf_val_pis)

        inserir_coluna_ao_lado(df_detalhe, 'VALOR COFINS', 'CONFERÊNCIA VALOR COFINS', conf_val_cofins)

        # Remover colunas auxiliares do merge para não sujar a planilha final
        df_detalhe.drop(columns=['CHAVE_LOOKUP', 'ICMS_G', 'IPI_G', 'PIS_G', 'CFOP_CLEAN'], inplace=True)

        st.success("✅ Processamento concluído com sucesso!")
        st.dataframe(df_detalhe.head(50))

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_detalhe.to_excel(writer, index=False, sheet_name='Detalhamento')
            df_grade.to_excel(writer, index=False, sheet_name='Grade')
        
        st.download_button(label="📥 Baixar Planilha Processada", data=output.getvalue(), file_name="Conferencia_Fiscal.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
