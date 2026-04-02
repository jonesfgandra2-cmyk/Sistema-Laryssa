import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Automação de Fechamento", layout="wide")

st.title("📊 Automação de Conferência ICMS/IPI")
st.markdown("Arraste sua planilha aqui para processar as colunas de conferência automaticamente.")

uploaded_file = st.file_uploader("Upload da Planilha Excel", type="xlsx")

if uploaded_file:
    try:
        # 1. Carregar as abas
        df_detalhe = pd.read_excel(uploaded_file, sheet_name='Detalhamento')
        df_grade = pd.read_excel(uploaded_file, sheet_name='Grade')

        st.info("Planilha carregada com sucesso! Processando regras...")

        # --- LÓGICA DO PROCX (XLOOKUP) ---
        # No PROCX: PROCX(A2; Grade!D:D; Grade!K:K)
        # Vamos assumir que a Coluna D da Grade chama-se 'CHAVE' 
        # e a Coluna K chama-se 'ALÍQUOTA ICMS' na aba Grade.
        
        # Preparando df_grade para o join (pegando apenas o necessário)
        # Ajuste os nomes das colunas abaixo se forem diferentes no seu arquivo real
        df_grade_resumo = df_grade.iloc[:, [3, 10, 17]] # Colunas D (3), K (10) e R (17)
        df_grade_resumo.columns = ['CHAVE_LOOKUP', 'ICMS_GRADE_VAL', 'IPI_GRADE_VAL']

        # Fazendo o De/Para (Merge)
        df_final = pd.merge(
            df_detalhe, 
            df_grade_resumo, 
            left_on='CHAVE', 
            right_on='CHAVE_LOOKUP', 
            how='left'
        )

        # --- CÁLCULOS ICMS ---
        # Inserindo ALÍQUOTA ICMS GRADE
        df_final['ALÍQUOTA ICMS GRADE'] = df_final['ICMS_GRADE_VAL']
        
        # CONFERÊNCIA ALÍQUOTA ICMS (AB == AC)
        # Nota: Ajustamos os nomes conforme a planilha (ex: 'ALÍQUOTA ICMS')
        df_final['CONFERÊNCIA ALÍQUOTA ICMS'] = df_final['ALÍQUOTA ICMS'] == df_final['ALÍQUOTA ICMS GRADE']
        
        # CONFERÊNCIA VALOR ICMS: (VALOR BASE * ALIQ %) - VALOR ICMS
        # Fórmula: (Z2 * AB2%) - AE2
        df_final['CONFERÊNCIA VALOR ICMS'] = (df_final['VALOR BASE ICMS'] * (df_final['ALÍQUOTA ICMS'] / 100)) - df_final['VALOR ICMS']

        # --- CÁLCULOS IPI ---
        # Inserindo ALIQUOTA IPI GRADE
        df_final['ALIQUOTA IPI GRADE'] = df_final['IPI_GRADE_VAL']
        
        # CONFERÊNCIA ALÍQUOTA IPI (AK == AL)
        df_final['CONFERÊNCIA ALÍQUOTA IPI'] = df_final['ALÍQUOTA IPI'] == df_final['ALIQUOTA IPI GRADE']
        
        # CONFERÊNCIA VALOR IPI: (VALOR BASE IPI * ALIQ IPI %) - VALOR IPI
        # Fórmula: (AI2 * AK2%) - AN2
        df_final['CONFERÊNCIA VALOR IPI'] = (df_final['VALOR BASE IPI'] * (df_final['ALÍQUOTA IPI'] / 100)) - df_final['VALOR IPI']

        # Remover coluna auxiliar de join
        df_final.drop(columns=['CHAVE_LOOKUP', 'ICMS_GRADE_VAL', 'IPI_GRADE_VAL'], inplace=True)

        # Reorganizar colunas (opcional: colocar as novas colunas ao lado das originais)
        # Para manter simples, as novas colunas irão para o final do arquivo.

        st.success("✅ Processamento concluído!")
        st.dataframe(df_final.head())

        # --- EXPORTAÇÃO ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Detalhamento_Processado')
            # Mantém a aba Grade original no arquivo final
            df_grade.to_excel(writer, index=False, sheet_name='Grade')
        
        st.download_button(
            label="📥 Baixar Planilha Conferida",
            data=output.getvalue(),
            file_name="Fechamento_Processado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
        st.warning("Verifique se os nomes das colunas (CHAVE, ALÍQUOTA ICMS, etc) estão idênticos aos da planilha.")
