import streamlit as st
import pandas as pd
from io import BytesIO
import time

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="FiscalFlow Pro | Auditoria Fiscal",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- 2. INICIALIZAÇÃO DE PARÂMETROS PADRÕES ---
if 'params' not in st.session_state:
    st.session_state.params = {
        'cfops_icms': "1301, 1302, 1303, 1352, 1551, 1556, 1905, 1908, 1911, 1916, 1920, 1933, 2551, 2556, 2911, 5152, 5502, 5551, 5906, 5908, 5909, 5911, 5921, 6551, 6908, 6911, 6915, 7102",
        'cfops_ipi': "1301, 1302, 1303, 1352, 1551, 1556, 1905, 1908, 1911, 1916, 1920, 1933, 2551, 2556, 2911, 5502, 5551, 5906, 5908, 5909, 5911, 5921, 6551, 6908, 6911, 6915, 7102",
        'cfops_pis': "1301, 1302, 1303, 1352, 1551, 1556, 1905, 1908, 1911, 1916, 1920, 1933, 2551, 2556, 2911, 5152, 5502, 5551, 5906, 5908, 5909, 5911, 5921, 6551, 6908, 6911, 6915, 7102, 5910, 6910",
        'idx_chave': 3,
        'idx_icms': 9,
        'idx_ipi': 16,
        'idx_pis': 18
    }

# --- 3. ESTILO CSS PARA CUSTOMIZAÇÃO ---
st.markdown("""
    <style>
    /* Estilo do título e sidebar */
    .main { background-color: #f4f7f6; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { 
        background-color: #e1e4e8; 
        border-radius: 5px 5px 0px 0px; 
        padding: 10px 20px;
        font-weight: bold;
    }
    .stTabs [aria-selected="true"] { background-color: #1E3A8A !important; color: white !important; }
    
    /* Logo e Header custom */
    .fiscal-header {
        background-color: #1E3A8A;
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin-bottom: 25px;
        display: flex;
        align-items: center;
        gap: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 4. SIDEBAR COM LOGO ---
with st.sidebar:
    # Logo Fiscal (Ícone de Balança/Documento)
    st.image("https://cdn-icons-png.flaticon.com/512/5833/5833864.png", width=100)
    st.title("FiscalFlow Pro")
    st.subheader("Auditoria de Alíquotas")
    st.markdown("---")
    st.info("Sistema inteligente de conferência de CFOPs e Alíquotas Grade.")
    
    if st.button("🔄 Resetar Sistema"):
        st.rerun()

# --- 5. CABEÇALHO PRINCIPAL ---
st.markdown("""
    <div class="fiscal-header">
        <img src="https://cdn-icons-png.flaticon.com/512/11516/11516492.png" width="60">
        <div>
            <h1 style="margin:0; color: white;">Módulo de Automação de Rotina</h1>
            <p style="margin:0; opacity: 0.8;">Gestão Tributária | Detalhamento x Grade</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

# --- 6. INTERFACE (TABS) ---
tab_proc, tab_param, tab_inst = st.tabs(["🚀 Processamento", "⚙️ Parâmetros Técnicos", "📖 Guia de Uso"])

# --- TAB: PARÂMETROS ---
with tab_param:
    st.subheader("🛠️ Configuração de Regras")
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.markdown("#### Listas de Exceção (CFOP)")
        st.caption("Insira os CFOPs separados por vírgula para forçar alíquota 0%")
        new_icms = st.text_area("ICMS Grade 0%", value=st.session_state.params['cfops_icms'], height=100)
        new_ipi = st.text_area("IPI Grade 0%", value=st.session_state.params['cfops_ipi'], height=100)
        new_pis = st.text_area("PIS Grade 0%", value=st.session_state.params['cfops_pis'], height=100)

    with col_b:
        st.markdown("#### Mapeamento de Colunas (Aba Grade)")
        st.caption("Defina a posição das colunas no Excel da Grade (0=A, 1=B...)")
        new_idx_chave = st.number_input("Índice CHAVE (Col. D=3)", value=st.session_state.params['idx_chave'])
        new_idx_icms = st.number_input("Índice ICMS (Col. J=9)", value=st.session_state.params['idx_icms'])
        new_idx_ipi = st.number_input("Índice IPI (Col. Q=16)", value=st.session_state.params['idx_ipi'])
        new_idx_pis = st.number_input("Índice PIS (Col. S=18)", value=st.session_state.params['idx_pis'])

    if st.button("💾 Salvar Configurações"):
        st.session_state.params.update({
            'cfops_icms': new_icms, 'cfops_ipi': new_ipi, 'cfops_pis': new_pis,
            'idx_chave': new_idx_chave, 'idx_icms': new_idx_icms,
            'idx_ipi': new_idx_ipi, 'idx_pis': new_idx_pis
        })
        st.success("Parâmetros salvos com sucesso!")

# --- TAB: PROCESSAMENTO ---
with tab_proc:
    uploaded_file = st.file_uploader("Selecione a Planilha (.xlsx)", type="xlsx", help="Upload do arquivo Detalhamento + Grade")

    if uploaded_file:
        try:
            with st.status("Auditando dados fiscais...", expanded=True) as status:
                xls = pd.ExcelFile(uploaded_file)
                aba_detalhe = [s for s in xls.sheet_names if 'detalhamento' in s.lower()][0]
                aba_grade = [s for s in xls.sheet_names if 'grade' in s.lower()][0]
                
                df_detalhe = pd.read_excel(uploaded_file, sheet_name=aba_detalhe, dtype=str)
                df_grade = pd.read_excel(uploaded_file, sheet_name=aba_grade, dtype=str)

                # Limpeza de CFOP
                def limpar(txt): return [c.strip() for c in txt.split(',')]
                l_icms = limpar(st.session_state.params['cfops_icms'])
                l_ipi = limpar(st.session_state.params['cfops_ipi'])
                l_pis = limpar(st.session_state.params['cfops_pis'])

                # Conversão Numérica
                cols_n = ['ALIQUOTA ICMS', 'BASE DE CALCULO ICMS', 'VALOR ICMS', 'ALIQUOTA IPI', 'BASE DE CALCULO IPI', 'VALOR IPI', 'ALIQUOTA PIS', 'BASE DE CALCULO PIS', 'VALOR PIS', 'BASE DE CALCULO COFINS', 'ALIQUOTA COFINS', 'VALOR COFINS']
                for c in cols_n:
                    if c in df_detalhe.columns:
                        df_detalhe[c] = pd.to_numeric(df_detalhe[c], errors='coerce').fillna(0)
                
                df_detalhe['CFOP_CLEAN'] = df_detalhe['CFOP'].astype(str).str.strip()

                # Merge Grade
                p = st.session_state.params
                df_g = df_grade.iloc[:, [p['idx_chave'], p['idx_icms'], p['idx_ipi'], p['idx_pis']]].copy()
                df_g.columns = ['C_LOOK', 'I_G', 'IP_G', 'P_G']
                df_g = df_g.drop_duplicates(subset=['C_LOOK'])
                for c in ['I_G', 'IP_G', 'P_G']: df_g[c] = pd.to_numeric(df_g[c], errors='coerce')

                df_detalhe = pd.merge(df_detalhe, df_g, left_on='CHAVE', right_on='C_LOOK', how='left')

                # Regras de Negócio
                def logic(df, col_g, lista):
                    res_num = df[col_g].copy()
                    res_num.loc[df['CFOP_CLEAN'].isin(lista)] = 0
                    res_c = res_num.fillna(0)
                    disp = res_num.astype(str)
                    disp.loc[df['C_LOOK'].isna() & ~df['CFOP_CLEAN'].isin(lista)] = "Não Encontrado"
                    return res_c, disp

                v_icms, d_icms = logic(df_detalhe, 'I_G', l_icms)
                v_ipi, d_ipi = logic(df_detalhe, 'IP_G', l_ipi)
                v_pis, d_pis = logic(df_detalhe, 'P_G', l_pis)

                # Cálculos Finais
                def inserir(ref, nome, val):
                    if nome in df_detalhe.columns: df_detalhe.drop(columns=[nome], inplace=True)
                    df_detalhe.insert(df_detalhe.columns.get_loc(ref)+1, nome, val)

                inserir('ALIQUOTA ICMS', 'ALÍQUOTA ICMS GRADE', d_icms)
                inserir('ALÍQUOTA ICMS GRADE', 'CONFERÊNCIA ALÍQUOTA ICMS', df_detalhe['ALIQUOTA ICMS'] == v_icms)
                inserir('VALOR ICMS', 'CONFERÊNCIA VALOR ICMS', (df_detalhe['BASE DE CALCULO ICMS'] * (df_detalhe['ALIQUOTA ICMS'] / 100)) - df_detalhe['VALOR ICMS'])
                
                inserir('ALIQUOTA IPI', 'ALIQUOTA IPI GRADE', d_ipi)
                inserir('ALIQUOTA IPI GRADE', 'CONFERÊNCIA ALÍQUOTA IPI', df_detalhe['ALIQUOTA IPI'] == v_ipi)
                inserir('VALOR IPI', 'CONFERÊNCIA VALOR IPI', (df_detalhe['BASE DE CALCULO IPI'] * (df_detalhe['ALIQUOTA IPI'] / 100)) - df_detalhe['VALOR IPI'])
                
                inserir('ALIQUOTA PIS', 'ALIQUOTA PIS GRADE', d_pis)
                inserir('ALIQUOTA PIS GRADE', 'CONFERÊNCIA ALÍQUOTA PIS', df_detalhe['ALIQUOTA PIS'] == v_pis)
                inserir('VALOR PIS', 'CONFERÊNCIA VALOR PIS', (df_detalhe['BASE DE CALCULO PIS'] * (df_detalhe['ALIQUOTA PIS'] / 100)) - df_detalhe['VALOR PIS'])
                
                inserir('VALOR COFINS', 'CONFERÊNCIA VALOR COFINS', (df_detalhe['BASE DE CALCULO COFINS'] * (df_detalhe['ALIQUOTA COFINS'] / 100)) - df_detalhe['VALOR COFINS'])

                df_detalhe.drop(columns=['C_LOOK', 'I_G', 'IP_G', 'P_G', 'CFOP_CLEAN'], inplace=True)
                status.update(label="Auditoria Concluída!", state="complete")

            # Dashboard
            st.subheader("📊 Resumo da Auditoria")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total de Itens", len(df_detalhe))
            c2.metric("Não Encontrados", (d_icms == "Não Encontrado").sum())
            c3.metric("Divergência ICMS", (df_detalhe['CONFERÊNCIA ALÍQUOTA ICMS'] == False).sum())
            c4.metric("Divergência IPI", (df_detalhe['CONFERÊNCIA ALÍQUOTA IPI'] == False).sum())

            st.dataframe(df_detalhe.head(50), use_container_width=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_detalhe.to_excel(writer, index=False, sheet_name='Detalhamento')
                df_grade.to_excel(writer, index=False, sheet_name='Grade')
            
            st.download_button(
                label="📥 Baixar Planilha Auditada",
                data=output.getvalue(),
                file_name=f"FiscalFlow_Auditoria_{time.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro no processamento: {e}")

# --- TAB: GUIA ---
with tab_inst:
    st.subheader("📖 Guia Rápido")
    st.markdown("""
    1. **Identidade Visual:** O sistema usa a coluna `CHAVE` para cruzar a Grade com o Detalhamento.
    2. **Parâmetros:** Na aba de parâmetros, você define quais CFOPs são considerados exceções (Alíquota 0%).
    3. **Indicadores:** O Dashboard mostra rapidamente onde estão os problemas (Divergências e Chaves Não Encontradas).
    4. **Segurança:** O arquivo original não é alterado, o sistema gera uma nova planilha auditada para download.
    """)
