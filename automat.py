import streamlit as st
import pandas as pd
from io import BytesIO
import time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="FiscalFlow | Automação",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- ESTILO CSS PERSONALIZADO ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #e0e0e0; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .stDownloadButton>button { background-color: #28a745 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# --- BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2618/2618245.png", width=100)
    st.title("FiscalFlow v2.0")
    st.info("Sistema de Auditoria e Conferência de Alíquotas ICMS, IPI, PIS e COFINS.")
    
    st.markdown("---")
    st.subheader("⚙️ Configurações")
    mostrar_preview = st.checkbox("Mostrar Preview dos Dados", value=True)
    
    with st.expander("📝 Regras Aplicadas"):
        st.caption("**Exceções CFOP:** 0% para ICMS/IPI/PIS em listas específicas.")
        st.caption("**Não Encontrado:** Sinaliza chaves ausentes na Grade.")
        st.caption("**Cálculo:** (Base * Alíquota%) - Valor.")

# --- TÍTULO PRINCIPAL ---
st.title("🚀 Automação de Rotina Fiscal")
st.markdown("Faça o upload da sua planilha para iniciar a conferência automática.")

# --- ÁREA DE UPLOAD ---
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type="xlsx")

# --- PROCESSAMENTO ---
if uploaded_file:
    try:
        with st.status("Processando dados...", expanded=True) as status:
            st.write("📖 Lendo abas Detalhamento e Grade...")
            xls = pd.ExcelFile(uploaded_file)
            aba_detalhe = [s for s in xls.sheet_names if 'detalhamento' in s.lower()][0]
            aba_grade = [s for s in xls.sheet_names if 'grade' in s.lower()][0]
            
            df_detalhe = pd.read_excel(uploaded_file, sheet_name=aba_detalhe, dtype=str)
            df_grade = pd.read_excel(uploaded_file, sheet_name=aba_grade, dtype=str)

            st.write("🛠️ Aplicando listas de exceção...")
            cfops_base = ['1301', '1302', '1303', '1352', '1551', '1556', '1905', '1908', '1911', '1916', '1920', '1933', '2551', '2556', '2911', '5502', '5551', '5906', '5908', '5909', '5911', '5921', '6551', '6908', '6911', '6915', '7102']
            cfops_icms_zero = cfops_base + ['5152']
            cfops_ipi_zero = cfops_base
            cfops_pis_zero = cfops_base + ['5152', '5910', '6910']

            # Conversão e Limpeza
            for col in ['ALIQUOTA ICMS', 'BASE DE CALCULO ICMS', 'VALOR ICMS', 'ALIQUOTA IPI', 'BASE DE CALCULO IPI', 'VALOR IPI', 'ALIQUOTA PIS', 'BASE DE CALCULO PIS', 'VALOR PIS', 'BASE DE CALCULO COFINS', 'ALIQUOTA COFINS', 'VALOR COFINS']:
                if col in df_detalhe.columns:
                    df_detalhe[col] = pd.to_numeric(df_detalhe[col], errors='coerce').fillna(0)
            
            df_detalhe['CFOP_CLEAN'] = df_detalhe['CFOP'].astype(str).str.strip()

            st.write("🔍 Cruzando dados com a Grade...")
            df_grade_resumo = df_grade.iloc[:, [3, 9, 16, 18]].copy()
            df_grade_resumo.columns = ['CHAVE_LOOKUP', 'ICMS_G', 'IPI_G', 'PIS_G']
            df_grade_resumo = df_grade_resumo.drop_duplicates(subset=['CHAVE_LOOKUP'])
            for col in ['ICMS_G', 'IPI_G', 'PIS_G']:
                df_grade_resumo[col] = pd.to_numeric(df_grade_resumo[col], errors='coerce')

            df_detalhe = pd.merge(df_detalhe, df_grade_resumo, left_on='CHAVE', right_on='CHAVE_LOOKUP', how='left')

            # Função de Regra
            def aplicar_logica(df, col_grade, lista):
                res_num = df[col_grade].copy()
                res_num.loc[df['CFOP_CLEAN'].isin(lista)] = 0
                res_clean = res_num.fillna(0)
                display = res_num.astype(str)
                display.loc[df['CHAVE_LOOKUP'].isna() & ~df['CFOP_CLEAN'].isin(lista)] = "Não Encontrado"
                return res_clean, display

            v_icms_n, d_icms = aplicar_logica(df_detalhe, 'ICMS_G', cfops_icms_zero)
            v_ipi_n, d_ipi = aplicar_logica(df_detalhe, 'IPI_G', cfops_ipi_zero)
            v_pis_n, d_pis = aplicar_logica(df_detalhe, 'PIS_G', cfops_pis_zero)

            st.write("🧮 Calculando divergências...")
            conf_val_icms = (df_detalhe['BASE DE CALCULO ICMS'] * (df_detalhe['ALIQUOTA ICMS'] / 100)) - df_detalhe['VALOR ICMS']
            conf_val_ipi = (df_detalhe['BASE DE CALCULO IPI'] * (df_detalhe['ALIQUOTA IPI'] / 100)) - df_detalhe['VALOR IPI']
            conf_val_pis = (df_detalhe['BASE DE CALCULO PIS'] * (df_detalhe['ALIQUOTA PIS'] / 100)) - df_detalhe['VALOR PIS']
            conf_val_cofins = (df_detalhe['BASE DE CALCULO COFINS'] * (df_detalhe['ALIQUOTA COFINS'] / 100)) - df_detalhe['VALOR COFINS']

            # Inserções
            def inserir(ref, nome, val):
                if nome in df_detalhe.columns: df_detalhe.drop(columns=[nome], inplace=True)
                idx = df_detalhe.columns.get_loc(ref) + 1
                df_detalhe.insert(idx, nome, val)

            inserir('ALIQUOTA ICMS', 'ALÍQUOTA ICMS GRADE', d_icms)
            inserir('ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', df_detalhe['ALIQUOTA ICMS'] == v_icms_n)
            inserir('VALOR ICMS', 'CONFERÊNCIA VALOR ICMS', conf_val_icms)
            inserir('ALIQUOTA IPI', 'ALIQUOTA IPI GRADE', d_ipi)
            inserir('ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', df_detalhe['ALIQUOTA IPI'] == v_ipi_n)
            inserir('VALOR IPI', 'CONFERÊNCIA VALOR IPI', conf_val_ipi)
            inserir('ALIQUOTA PIS', 'ALIQUOTA PIS GRADE', d_pis)
            inserir('ALIQUOTA PIS GRADE', 'CONFERÊNCIA ALÍQUOTA PIS', df_detalhe['ALIQUOTA PIS'] == v_pis_n)
            inserir('VALOR PIS', 'CONFERÊNCIA VALOR PIS', conf_val_pis)
            inserir('VALOR COFINS', 'CONFERÊNCIA VALOR COFINS', conf_val_cofins)

            df_detalhe.drop(columns=['CHAVE_LOOKUP', 'ICMS_G', 'IPI_G', 'PIS_G', 'CFOP_CLEAN'], inplace=True)
            status.update(label="✅ Processamento concluído!", state="complete", expanded=False)

        # --- DASHBOARD DE MÉTRICAS ---
        st.markdown("### 📈 Resumo do Processamento")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total de Linhas", len(df_detalhe))
        m2.metric("Não Encontrados (ICMS)", (d_icms == "Não Encontrado").sum())
        m3.metric("Divergências ICMS", (df_detalhe['CONFERÊNCIA ALÍQUOTA ICMS'] == False).sum())
        m4.metric("Divergências IPI", (df_detalhe['CONFERÊNCIA ALÍQUOTA IPI'] == False).sum())

        # --- VISUALIZAÇÃO E DOWNLOAD ---
        tab1, tab2 = st.tabs(["📄 Visualizar Resultado", "📥 Download"])
        
        with tab1:
            if mostrar_preview:
                st.dataframe(df_detalhe.head(100), use_container_width=True)
        
        with tab2:
            st.success("O arquivo está pronto para ser baixado. Clique no botão abaixo.")
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_detalhe.to_excel(writer, index=False, sheet_name='Detalhamento')
                df_grade.to_excel(writer, index=False, sheet_name='Grade')
            
            st.download_button(
                label="💾 Baixar Planilha Processada (.xlsx)",
                data=output.getvalue(),
                file_name=f"Conferencia_Fiscal_{time.strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Erro Crítico: {e}")
        st.warning("Certifique-se de que as abas 'Detalhamento' e 'Grade' existam e as colunas estejam corretas.")
else:
    st.info("👆 Aguardando upload do arquivo Excel para iniciar.")
