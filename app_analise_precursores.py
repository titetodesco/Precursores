import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
import squarify
import pdfplumber
import docx
from fuzzywuzzy import fuzz
import unicodedata
from langdetect import detect
import os
import io

# ---- Funções auxiliares ----

@st.cache_data
def load_precursors_excel(path):
    df = pd.read_excel(path)
    if not {"Dimensao", "PT", "EN"}.issubset(df.columns):
        st.error("A planilha deve conter as colunas: 'Dimensao', 'PT', 'EN'.")
        st.stop()
    return df.dropna(subset=["PT", "EN"])

def extract_text_from_pdf(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        text = " ".join(page.extract_text() for page in pdf.pages if page.extract_text())
    return text

def extract_text_from_docx(uploaded_file):
    doc = docx.Document(uploaded_file)
    return " ".join([para.text for para in doc.paragraphs])

def normalize(text):
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join([c for c in text if not unicodedata.combining(c)])
    return text.lower()

def fuzzy_match_terms_count(text, precursors_df, threshold=75):
    import re
    normalized_text = normalize(text)
    results = []
    for _, row in precursors_df.iterrows():
        for lang in ["PT", "EN"]:
            for term in str(row[lang]).split(";"):
                term = normalize(term.strip())
                # Conta quantas vezes o termo ocorre de forma aproximada no texto
                count = 0
                for sentence in re.split(r'[.!?]', normalized_text):
                    if fuzz.partial_ratio(term, sentence.strip()) >= threshold:
                        count += 1
                if count > 0:
                    results.append({
                        "Precursor": term,
                        "Dimensao": row["Dimensao"],
                        "Idioma": lang,
                        "Frequência": count
                    })
    return pd.DataFrame(results)


# ---- App Streamlit ----

st.set_page_config(page_title="Análise de Precursores em Relatórios de Acidente", layout="wide")
st.title("🔎 Análise de Precursores em Relatórios de Acidente")

st.sidebar.markdown("### 1️⃣ Upload do relatório")
uploaded_report = st.sidebar.file_uploader("Selecione o relatório de acidente (.pdf, .docx)", type=["pdf", "docx"])
threshold = st.sidebar.slider("Threshold de similaridade fuzzy (%)", min_value=60, max_value=100, value=75)
st.sidebar.markdown("### 2️⃣ Base de precursores")
st.sidebar.info("A planilha de precursores deve estar disponível no GitHub.")

# Carrega precursores do repositório (ajuste o caminho se for diferente)
PRECURSOR_PATH = "precursores.xlsx"
try:
    PRECURSOR_PATH = "https://raw.githubusercontent.com/titetodesco/Precursores/main/precursores.xlsx"
    precursors_df = load_precursors_excel(PRECURSOR_PATH)
except Exception as e:
    st.error(f"Erro ao carregar precursores: {e}")
    st.stop()

if uploaded_report:
    # Extrai texto
    ext = os.path.splitext(uploaded_report.name)[-1].lower()
    if ext == ".pdf":
        text = extract_text_from_pdf(uploaded_report)
    elif ext == ".docx":
        text = extract_text_from_docx(uploaded_report)
    else:
        st.error("Formato de arquivo não suportado.")
        st.stop()

    st.markdown("#### 🔹 Exemplo do texto extraído")
    st.code(text[:300] + "..." if len(text) > 300 else text)

    # Detecta idioma principal
    try:
        idioma = detect(text)
        lang_detected = "EN" if idioma == "en" else "PT"
    except Exception:
        lang_detected = "PT"

    st.markdown(f"**Idioma detectado:** {lang_detected}")

    # Faz o matching
    st.info("Analisando o relatório, aguarde alguns segundos...")
    resultado = fuzzy_match_terms_count(text, precursors_df, threshold=threshold)

    if resultado.empty:
        st.warning("⚠️ Nenhum precursor foi identificado no relatório.")
    else:
        # Tabela detalhada de precursores encontrados
        resultado = resultado[resultado["Idioma"] == lang_detected]
        resumo = resultado.groupby(["Dimensao", "Precursor"])["Frequência"].sum().reset_index()
        st.success(f"🔍 Foram identificados {len(resumo)} precursores únicos.")
        st.dataframe(resumo)

        # Gráfico de barras
        df_freq = resumo.groupby("Dimensao")["Frequência"].sum().reset_index()
        fig1, ax1 = plt.subplots()
        sns.barplot(x="Dimensao", y="Frequência", data=df_freq, ax=ax1)
        ax1.set_title("Frequência de Precursores por Dimensão")
        st.pyplot(fig1)


        # Mapa de Árvore (Treemap)
        st.subheader("🌳 Mapa de Árvore - Precursores por Dimensão")
        fig2 = px.treemap(resumo, path=["Dimensao", "Precursor"], values="Frequência")
        st.plotly_chart(fig2, use_container_width=True)

        # Sunburst (igual à sua imagem)
        st.subheader("☀️ Distribuição de Precursores por Dimensão")
        fig3 = px.sunburst(resumo, path=["Dimensao", "Precursor"], values="Frequência", title="Distribuição de Precursores por Dimensão")
        st.plotly_chart(fig3, use_container_width=True)

        # Gera planilha Sim/Não (para o idioma detectado)
        # Usa diretamente os precursores encontrados no resultado do idioma
        encontrados_norm = resultado["Precursor"].str.lower().str.strip().unique().tolist()
        status_list = []
        for _, row in precursors_df.iterrows():
            for term in str(row[lang_detected]).split(";"):
                term_norm = term.strip().lower()
                status_list.append({
                    "Dimensao": row["Dimensao"],
                    "Idioma": lang_detected,
                    "Precursor": term.strip(),
                    "Encontrado": "Sim" if term_norm in encontrados_norm else "Não"
                })
        df_status = pd.DataFrame(status_list)

         
        # ====== Downloads em Excel (.xlsx) ======
        st.markdown("#### 📥 Baixar resultados em Excel")
        
        # 1. Download do resumo dos precursores encontrados
        output_resumo = io.BytesIO()
        with pd.ExcelWriter(output_resumo, engine='xlsxwriter') as writer:
            resumo.to_excel(writer, index=False, sheet_name='Resumo')
            output_resumo.seek(0)
            st.download_button(label="Baixar resumo (Excel)", data=output_resumo, file_name="precursores_resumo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # 2. Download da planilha Sim/Não
        output_status = io.BytesIO()
        with pd.ExcelWriter(output_status, engine='xlsxwriter') as writer:
            df_status.to_excel(writer, index=False, sheet_name='Status')
            output_status.seek(0)
            st.download_button(label="Baixar planilha Sim/Não (Excel)", data=output_status, file_name="status_precursores.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Faça upload do relatório (.pdf ou .docx) para iniciar a análise.")

# Rodapé
st.markdown("---")
st.markdown("App desenvolvido por Tite Todesco. Baseado no dicionário de precursores HTO.")

