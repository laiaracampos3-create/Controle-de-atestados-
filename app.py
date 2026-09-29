from datetime import datetime, timedelta
import json
import sqlite3
from google import genai
from google.genai import types
import pandas as pd
from pypdf import PdfReader
import streamlit as st

# --- Configuração da Página ---
st.set_page_config(
    page_title="Controle de Atestados - DP", page_icon="📋", layout="wide"
)

# Inicializa o cliente da IA (pega a chave salva nas variáveis de ambiente do Render)
client = genai.Client()


# --- Funções do Banco SQLite ---
def init_db():
    conn = sqlite3.connect("atestados.db")
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS atestados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula TEXT,
            colaborador TEXT NOT NULL,
            setor TEXT,
            tipo_atestado TEXT,
            data_inicio DATE NOT NULL,
            dias INTEGER NOT NULL,
            data_fim DATE NOT NULL,
            cid TEXT,
            impacto_inss TEXT,
            observacoes TEXT
        )
    """
    )
    conn.commit()
    conn.close()


def salvar_atestado(
    matricula, colab, setor, tipo, data_ini, dias, cid, obs
):
    data_fim = data_ini + timedelta(days=int(dias) - 1)
    inss = "Sim" if int(dias) > 15 else "Não"

    conn = sqlite3.connect("atestados.db")
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO atestados (matricula, colaborador, setor, tipo_atestado, data_inicio, dias, data_fim, cid, impacto_inss, observacoes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        (
            matricula,
            colab,
            setor,
            tipo,
            data_ini,
            dias,
            data_fim,
            cid,
            inss,
            obs,
        ),
    )
    conn.commit()
    conn.close()


def carregar_dados():
    conn = sqlite3.connect("atestados.db")
    df = pd.read_sql_query("SELECT * FROM atestados", conn)
    conn.close()
    if not df.empty:
        df["data_inicio"] = pd.to_datetime(df["data_inicio"]).dt.date
        df["data_fim"] = pd.to_datetime(df["data_fim"]).dt.date
    return df


init_db()


# --- Leitura do PDF ---
def extrair_dados_pdf(arquivo_pdf):
    reader = PdfReader(arquivo_pdf)
    texto = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            texto += t + "\n"

    prompt = """
    Extraia do atestado em formato JSON estrito:
    {
        "nome": "Nome completo do colaborador/paciente (ou vazio se não achar)",
        "dias": número inteiro de dias de afastamento (ex: 24h = 1, 48h = 2),
        "data_inicio": "AAAA-MM-DD",
        "cid": "Código do CID ou vazio"
    }
    """
    try:
        if len(texto.strip()) > 30:
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[prompt, texto],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json"
                ),
            )
        else:
            arquivo_pdf.seek(0)
            bytes_pdf = arquivo_pdf.read()
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[
                    types.Part.from_bytes(
                        data=bytes_pdf, mime_type="application/pdf"
                    ),
                    prompt,
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json"
                ),
            )
        return json.loads(resp.text)
    except Exception:
        return {}


# --- Interface Visual ---
st.title("📋 Painel de Controle de Atestados")

tab_upload, tab_fechamento, tab_historico = st.tabs(
    ["➕ Novo Atestado (Upload)", "📊 Espelho de Folha", "📂 Base de Dados"]
)

with tab_upload:
    st.subheader("1. Envie o PDF do Atestado")
    uploaded_file = st.file_uploader(
        "Selecione o arquivo digital ou foto/scan em PDF", type=["pdf"]
    )

    dados_iniciais = {}
    if uploaded_file is not None:
        if "dados_lidos" not in st.session_state:
            with st.spinner("Lendo documento com IA..."):
                st.session_state.dados_lidos = extrair_dados_pdf(
                    uploaded_file
                )
        dados_iniciais = st.session_state.dados_lidos
    else:
        st.session_state.pop("dados_lidos", None)

    st.markdown("---")
    st.subheader("2. Confirme os Dados do Lançamento")

    with st.form("form_cadastro", clear_on_submit=False):
        c1, c2, c3 = st.columns(3)
        matricula = c1.text_input("Matrícula")
        nome = c2.text_input(
            "Nome do Colaborador", value=dados_iniciais.get("nome", "")
        )
        setor = c3.text_input("Setor / Lotação")

        c4, c5, c6 = st.columns(3)
        tipo = c4.selectbox(
            "Tipo",
            [
                "Médico (Dias)",
                "Declaração de Comparecimento",
                "Acidente de Trabalho",
                "Acompanhamento",
            ],
        )

        data_padrao = datetime.today()
        if dados_iniciais.get("data_inicio"):
            try:
                data_padrao = datetime.strptime(
                    dados_iniciais["data_inicio"], "%Y-%m-%d"
                ).date()
            except Exception:
                pass

        data_ini = c5.date_input("Data de Início", value=data_padrao)
        dias = c6.number_input(
            "Qtd. Dias Afastados",
            value=int(dados_iniciais.get("dias") or 1),
            min_value=1,
            step=1,
        )

        c7, c8 = st.columns([1, 2])
        cid = c7.text_input("CID", value=dados_iniciais.get("cid", ""))
        obs = c8.text_input("Observações")

        gravar = st.form_submit_button("Salvar no Sistema")
        if gravar:
            if nome:
                salvar_atestado(
                    matricula,
                    nome,
                    setor,
                    tipo,
                    data_ini,
                    dias,
                    cid,
                    obs,
                )
                st.success(f"Atestado de {nome} salvo com sucesso!")
            else:
                st.error("Informe ao menos o nome do colaborador.")

with tab_fechamento:
    st.subheader("Corte da Competência para Fechamento")
    df = carregar_dados()

    c_dt1, c_dt2 = st.columns(2)
    dt_corte_inicio = c_dt1.date_input(
        "Início da Apuração", datetime.today().replace(day=1)
    )
    dt_corte_fim = c_dt2.date_input(
        "Fim da Apuração", datetime.today()
    )

    if not df.empty:
        mask = (df["data_inicio"] <= dt_corte_fim) & (
            df["data_fim"] >= dt_corte_inicio
        )
        df_periodo = df[mask].copy()

        if not df_periodo.empty:

            def calc_dias_comp(row):
                ini = max(row["data_inicio"], dt_corte_inicio)
                fim = min(row["data_fim"], dt_corte_fim)
                return (fim - ini).days + 1

            df_periodo["dias_na_competencia"] = df_periodo.apply(
                calc_dias_comp, axis=1
            )

            # KPIs
            k1, k2, k3 = st.columns(3)
            k1.metric("Atestados no Período", len(df_periodo))
            k2.metric(
                "Dias a Lançar na Folha",
                int(df_periodo["dias_na_competencia"].sum()),
            )
            k3.metric(
                "Atenção INSS (>15d)",
                len(df_periodo[df_periodo["impacto_inss"] == "Sim"]),
            )

            # Agrupamento por colaborador para bater ponto/folha
            resumo = (
                df_periodo.groupby(
                    ["matricula", "colaborador", "setor"],
                    as_index=False,
                )["dias_na_competencia"]
                .sum()
                .rename(
                    columns={"dias_na_competencia": "Total Dias a Abonar"}
                )
            )
            st.dataframe(resumo, use_container_width=True)

            csv = resumo.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "📥 Exportar Planilha de Fechamento (CSV)",
                data=csv,
                file_name=f"fechamento_folha_{dt_corte_inicio}_a_{dt_corte_fim}.csv",
                mime="text/csv",
            )
        else:
            st.info("Nenhum afastamento no período selecionado.")
    else:
        st.info("Nenhum dado cadastrado.")

with tab_historico:
    st.subheader("Histórico Completo")
    df = carregar_dados()
    if not df.empty:
        st.dataframe(df, use_container_width=True)
    else:
        st.write("Vazio.")
