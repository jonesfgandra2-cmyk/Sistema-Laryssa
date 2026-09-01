import streamlit as st
import pandas as pd
import io
import copy

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="SPED · Corretor GTIN",
    page_icon="🔖",
    layout="wide",
)

# ─────────────────────────────────────────────
# STYLE
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Remove default Streamlit padding */
.block-container { padding-top: 2rem; padding-bottom: 2rem; }

/* Header strip */
.header-strip {
    background: #0F1117;
    border-left: 4px solid #00C896;
    padding: 1.4rem 1.8rem;
    border-radius: 6px;
    margin-bottom: 2rem;
}
.header-strip h1 {
    color: #FFFFFF;
    font-size: 1.5rem;
    font-weight: 700;
    margin: 0 0 0.2rem 0;
    letter-spacing: -0.3px;
}
.header-strip p {
    color: #9CA3AF;
    font-size: 0.85rem;
    margin: 0;
}

/* Upload cards */
.upload-label {
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #6B7280;
    margin-bottom: 0.4rem;
}

/* Stat cards */
.stat-row {
    display: flex;
    gap: 1rem;
    margin: 1.5rem 0;
}
.stat-card {
    flex: 1;
    background: #F9FAFB;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    padding: 1rem 1.2rem;
}
.stat-card.red   { border-left: 3px solid #EF4444; }
.stat-card.green { border-left: 3px solid #10B981; }
.stat-card.blue  { border-left: 3px solid #3B82F6; }
.stat-card.gray  { border-left: 3px solid #9CA3AF; }
.stat-num {
    font-size: 1.8rem;
    font-weight: 700;
    color: #111827;
    line-height: 1;
}
.stat-label {
    font-size: 0.75rem;
    color: #6B7280;
    margin-top: 0.3rem;
}

/* Table tweaks */
.stDataFrame { border-radius: 8px; overflow: hidden; }

/* Step badge */
.step-badge {
    display: inline-block;
    background: #00C896;
    color: #fff;
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    padding: 2px 8px;
    border-radius: 20px;
    margin-bottom: 0.5rem;
    text-transform: uppercase;
}

/* Mono code look */
.mono { font-family: 'JetBrains Mono', monospace; font-size: 0.8rem; }

/* Download button override */
.stDownloadButton > button {
    background: #0F1117 !important;
    color: #00C896 !important;
    border: 1.5px solid #00C896 !important;
    font-weight: 600 !important;
    border-radius: 6px !important;
    padding: 0.5rem 1.4rem !important;
}
.stDownloadButton > button:hover {
    background: #00C896 !important;
    color: #0F1117 !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="header-strip">
    <h1>🔖 SPED Fiscal · Corretor de GTIN</h1>
    <p>Corrige automaticamente o campo <span class="mono">COD_BARRA</span> no registro <span class="mono">|0220|</span>
    com base na Planilha GTIN (Coluna D = CÓD 0220).</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CORE FUNCTIONS
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_planilha(file_bytes: bytes) -> dict:
    """
    Reads the GTIN spreadsheet and returns a dict:
    { 'V769': '27898637355887', ... }
    Keys are the COD_ITEM (col A), value is COD_0220 (col D).
    Only rows where col D is not empty are included.
    """
    df = pd.read_excel(io.BytesIO(file_bytes), header=0)
    # Columns: A=COD_ITEM, B=DESCRIÇÃO, C=CÓD 0200, D=CÓD 0220
    col_item = df.columns[0]
    col_0220 = df.columns[3]

    lookup = {}
    for _, row in df.iterrows():
        cod = str(row[col_item]).strip() if pd.notna(row[col_item]) else ""
        val = str(row[col_0220]).strip() if pd.notna(row[col_0220]) else ""
        # Remove .0 suffix that Excel sometimes adds
        if val.endswith(".0"):
            val = val[:-2]
        if cod and val and val not in ("nan", ""):
            lookup[cod] = val
    return lookup


def parse_sped(content: str):
    """
    Parses SPED text. Returns list of lines and a structure:
    { cod_item: { 'key': first4, '0220_barra': str } }
    tracking which 0220 COD_BARRA belongs to which item.
    """
    lines = content.splitlines(keepends=True)
    current_item = None
    item_map = {}  # line_index -> (cod_item, key4)

    for idx, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("|0200|"):
            parts = stripped.split("|")
            raw = parts[2].strip()
            current_item = raw.split()[0] if raw else raw
        elif stripped.startswith("|0220|") and current_item:
            item_map[idx] = current_item

    return lines, item_map


def analyse(lines, item_map, planilha):
    """
    Returns list of dicts with all 0220 records found, flagging discrepancies.
    """
    results = []
    for idx, cod_item in item_map.items():
        key4 = cod_item[:4]
        parts = lines[idx].strip().split("|")
        atual = parts[4] if len(parts) > 4 else ""
        esperado = planilha.get(key4, None)

        results.append({
            "linha": idx + 1,
            "cod_item": cod_item,
            "barra_atual": atual,
            "barra_esperada": esperado,
            "diverge": esperado is not None and atual != esperado,
            "sem_planilha": esperado is None,
        })
    return results


def apply_corrections(lines, item_map, planilha):
    """
    Returns corrected lines list and count of changes made.
    """
    new_lines = list(lines)
    count = 0
    for idx, cod_item in item_map.items():
        key4 = cod_item[:4]
        if key4 not in planilha:
            continue
        expected = planilha[key4]
        parts = new_lines[idx].rstrip("\r\n").split("|")
        atual = parts[4] if len(parts) > 4 else ""
        if atual != expected:
            parts[4] = expected
            ending = "\r\n" if new_lines[idx].endswith("\r\n") else "\n"
            new_lines[idx] = "|".join(parts) + ending
            count += 1
    return new_lines, count


# ─────────────────────────────────────────────
# STEP 1 — UPLOAD
# ─────────────────────────────────────────────
st.markdown('<div class="step-badge">Passo 1</div>', unsafe_allow_html=True)
st.markdown("#### Envie os arquivos")

col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="upload-label">Arquivo SPED Fiscal (.txt)</div>', unsafe_allow_html=True)
    sped_file = st.file_uploader("SPED", type=["txt"], label_visibility="collapsed", key="sped")

with col2:
    st.markdown('<div class="upload-label">Planilha GTIN (.xlsx)</div>', unsafe_allow_html=True)
    gtin_file = st.file_uploader("GTIN", type=["xlsx"], label_visibility="collapsed", key="gtin")

# ─────────────────────────────────────────────
# PROCESSING
# ─────────────────────────────────────────────
if sped_file and gtin_file:

    with st.spinner("Lendo Planilha GTIN..."):
        planilha = load_planilha(gtin_file.read())

    with st.spinner("Analisando arquivo SPED..."):
        raw_bytes = sped_file.read()
        # Try common encodings
        for enc in ("utf-8", "latin-1", "cp1252"):
            try:
                content = raw_bytes.decode(enc)
                used_encoding = enc
                break
            except UnicodeDecodeError:
                continue

        lines, item_map = parse_sped(content)
        results = analyse(lines, item_map, planilha)

    divergencias = [r for r in results if r["diverge"]]
    corretos     = [r for r in results if not r["diverge"] and not r["sem_planilha"]]
    sem_ref      = [r for r in results if r["sem_planilha"]]

    # ── STATS ────────────────────────────────
    st.markdown("---")
    st.markdown('<div class="step-badge">Passo 2</div>', unsafe_allow_html=True)
    st.markdown("#### Resultado da análise")

    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-card red">
            <div class="stat-num">{len(divergencias)}</div>
            <div class="stat-label">Divergências encontradas</div>
        </div>
        <div class="stat-card green">
            <div class="stat-num">{len(corretos)}</div>
            <div class="stat-label">Registros já corretos</div>
        </div>
        <div class="stat-card blue">
            <div class="stat-num">{len(item_map)}</div>
            <div class="stat-label">Total de registros |0220|</div>
        </div>
        <div class="stat-card gray">
            <div class="stat-num">{len(sem_ref)}</div>
            <div class="stat-label">Sem ref. na Planilha GTIN</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── DIVERGENCE TABLE ─────────────────────
    if divergencias:
        st.markdown("##### 🔴 Divergências detectadas")
        df_div = pd.DataFrame([{
            "Linha": r["linha"],
            "COD_ITEM": r["cod_item"],
            "COD_BARRA atual (SPED)": r["barra_atual"] if r["barra_atual"] else "*(vazio)*",
            "COD_BARRA esperado (Planilha)": r["barra_esperada"],
        } for r in divergencias])
        st.dataframe(df_div, use_container_width=True, hide_index=True)
    else:
        st.success("✅ Nenhuma divergência encontrada! O arquivo SPED já está correto.")

    if sem_ref:
        with st.expander(f"⚠️ {len(sem_ref)} itens sem referência na Planilha GTIN (mantidos sem alteração)"):
            df_sem = pd.DataFrame([{
                "Linha": r["linha"],
                "COD_ITEM": r["cod_item"],
                "COD_BARRA atual": r["barra_atual"],
            } for r in sem_ref])
            st.dataframe(df_sem, use_container_width=True, hide_index=True)

    # ── APPLY & DOWNLOAD ─────────────────────
    if divergencias:
        st.markdown("---")
        st.markdown('<div class="step-badge">Passo 3</div>', unsafe_allow_html=True)
        st.markdown("#### Aplicar correções e baixar")

        st.info(
            f"Serão corrigidos **{len(divergencias)} registros** no campo `COD_BARRA` do registro `|0220|`. "
            "Nenhum outro campo ou registro será alterado."
        )

        corrected_lines, n_fixed = apply_corrections(lines, item_map, planilha)
        corrected_content = "".join(corrected_lines)
        corrected_bytes = corrected_content.encode(used_encoding)

        original_name = sped_file.name.rsplit(".", 1)[0]
        output_name = f"{original_name}_corrigido.txt"

        st.download_button(
            label=f"⬇ Baixar SPED corrigido ({n_fixed} correções aplicadas)",
            data=corrected_bytes,
            file_name=output_name,
            mime="text/plain",
        )

        st.caption(f"Arquivo codificado em `{used_encoding}` · {len(corrected_lines)} linhas · {len(corrected_bytes):,} bytes")

else:
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; padding: 2.5rem; color: #9CA3AF; font-size: 0.9rem;">
        📂 Envie os dois arquivos acima para iniciar a análise.
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="font-size:0.75rem; color:#9CA3AF; text-align:center;">
    Regra aplicada: <span class="mono">|0220| COD_BARRA</span> deve ser igual ao valor da <b>Coluna D (CÓD 0220)</b>
    da Planilha GTIN, buscando pelo <b>COD_ITEM</b> (4 primeiros caracteres) do registro <span class="mono">|0200|</span> correspondente.
</div>
""", unsafe_allow_html=True)
