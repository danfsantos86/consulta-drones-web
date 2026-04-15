import os
import tempfile
from io import BytesIO

import pandas as pd
import streamlit as st

import extrair_drones


st.set_page_config(
    page_title="Consulta Drones Aprovados Anatel",
    page_icon="📄",
    layout="wide"
)


# =========================
# CSS premium / responsivo
# =========================
st.markdown("""
<style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }

    .header-wrap {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
        margin-top: 0.5rem;
        margin-bottom: 1.8rem;
    }

    .header-logo {
        display: flex;
        justify-content: center;
        margin-bottom: 0.8rem;
    }

    .header-title {
        text-align: center;
        font-size: 3rem;
        font-weight: 800;
        line-height: 1.1;
        margin: 0;
        padding: 0;
    }

    .header-subtitle {
        text-align: center;
        font-size: 1.05rem;
        color: #aeb6c2;
        margin-top: 0.75rem;
        margin-bottom: 0;
    }

    .section-divider {
        margin-top: 1.2rem;
        margin-bottom: 1.4rem;
    }

    .stMetric {
        background: rgba(255,255,255,0.02);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 14px;
        padding: 14px 16px;
    }

    @media (max-width: 768px) {
        .block-container {
            padding-top: 1.2rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }

        .header-title {
            font-size: 2.45rem;
        }

        .header-subtitle {
            font-size: 1rem;
            line-height: 1.5;
            max-width: 95%;
            margin-left: auto;
            margin-right: auto;
        }
    }
</style>
""", unsafe_allow_html=True)


def carregar_logo():
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
    if os.path.exists(logo_path):
        import base64

        with open(logo_path, "rb") as img_file:
            img_base64 = base64.b64encode(img_file.read()).decode()

        st.markdown(
            f"""
            <div style="display:flex; justify-content:center; align-items:center; margin-bottom: 12px;">
                <img src="data:image/png;base64,{img_base64}" width="140">
            </div>
            """,
            unsafe_allow_html=True
        )


def processar_docx_upload(uploaded_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded_file.getbuffer())
        caminho_temp = tmp.name

    try:
        extrair_drones.ARQUIVO_DOCX = caminho_temp
        registros = extrair_drones.carregar_drones()
        return registros
    finally:
        if os.path.exists(caminho_temp):
            os.remove(caminho_temp)


def filtrar_dataframe(df: pd.DataFrame, termo: str) -> pd.DataFrame:
    if not termo:
        return df

    termo = termo.lower()

    mask = (
        df["FABRICANTE"].astype(str).str.lower().str.contains(termo, na=False) |
        df["MODELO"].astype(str).str.lower().str.contains(termo, na=False) |
        df["NOME COMERCIAL"].astype(str).str.lower().str.contains(termo, na=False)
    )
    return df[mask]


def gerar_excel_em_memoria(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Drones")
    output.seek(0)
    return output


# =========================
# Header premium centralizado
# =========================
st.markdown('<div class="header-wrap">', unsafe_allow_html=True)
carregar_logo()
st.markdown(
    """
    <h1 class="header-title">Consulta Drones Aprovados Anatel</h1>
    <p class="header-subtitle">
        Versão web para consulta de drones aprovados pela Anatel
    </p>
    """,
    unsafe_allow_html=True
)
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.divider()

# =========================
# Upload
# =========================
uploaded_file = st.file_uploader(
    "Selecione o arquivo DOCX da Anatel",
    type=["docx"]
)

if uploaded_file is None:
    st.info("Faça o upload de um arquivo DOCX para iniciar a consulta.")
    st.stop()

# =========================
# Processamento
# =========================
try:
    registros = processar_docx_upload(uploaded_file)
except Exception as e:
    st.error(f"Erro ao processar o arquivo DOCX: {e}")
    st.stop()

if not registros:
    st.warning("Nenhum registro foi encontrado no arquivo enviado.")
    st.stop()

df = pd.DataFrame(registros)

# =========================
# Busca
# =========================
busca = st.text_input(
    "Busca global",
    placeholder="Digite fabricante, modelo ou nome comercial"
)

df_filtrado = filtrar_dataframe(df, busca)

# =========================
# Métricas
# =========================
col1, col2 = st.columns(2)

with col1:
    st.metric("Total de registros", len(df))

with col2:
    st.metric("Resultados filtrados", len(df_filtrado))

st.markdown("---")

# =========================
# Tabela
# =========================
st.dataframe(
    df_filtrado,
    width="stretch",
    hide_index=True
)

st.markdown("")

# =========================
# Download Excel
# =========================
excel_bytes = gerar_excel_em_memoria(df_filtrado)

st.download_button(
    label="📥 Baixar resultados em Excel",
    data=excel_bytes,
    file_name="drones_filtrados.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)