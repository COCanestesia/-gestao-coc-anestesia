import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
import gspread
from google.oauth2.service_account import Credentials

# ==========================================================
# CONFIGURAÇÃO INICIAL
# ==========================================================
st.set_page_config(page_title="Gestão COC Anestesia", layout="wide")
st.title("Registro de Cirurgias e Dashboard Financeiro")

# ==========================================================
# CONEXÃO GOOGLE SHEETS
# ==========================================================
@st.cache_resource
def conectar_sheets():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )

    return gspread.authorize(credentials)

try:
    client = conectar_sheets()
    URL_PLANILHA = "https://docs.google.com/spreadsheets/d/15J8J3dCVdGmn5GY8CWXef-JrpkxHcdX2CErKgWS0q7k/edit"
    planilha = client.open_by_url(URL_PLANILHA)

    aba_cirurgias = planilha.worksheet("CIRURGIAS")
    aba_convenios = planilha.worksheet("Página2")
    aba_cbhpm = planilha.worksheet("Página3")

except Exception as e:
    st.error(f"Erro de conexão com as planilhas: {e}")
    st.stop()

# ==========================================================
# CACHE DE DADOS
# ==========================================================
@st.cache_data(ttl=300)
def carregar_dados():
    df_cirurgias = pd.DataFrame(aba_cirurgias.get_all_records())
    df_convenios = pd.DataFrame(aba_convenios.get_all_records())
    df_cbhpm = pd.DataFrame(aba_cbhpm.get_all_records())
    return df_cirurgias, df_convenios, df_cbhpm

df_cirurgias, df_convenios, df_cbhpm = carregar_dados()

# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================
def limpar_moeda(valor):
    if pd.isna(valor):
        return 0.0

    valor_str = (
        str(valor)
        .replace("R$", "")
        .replace(".", "")
        .replace(",", ".")
        .strip()
    )

    try:
        return float(valor_str)
    except ValueError:
        return 0.0


def formatar_real(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


def converter_para_horas_robusto(coluna):
    return (
        pd.to_timedelta(coluna, errors="coerce")
        .dt.total_seconds() / 3600
    )


# ==========================================================
# MAPAS BASE
# ==========================================================
mapa_cbhpm = (
    df_cbhpm.set_index("Código").to_dict("index")
    if not df_cbhpm.empty else {}
)

mapa_convenios = (
    df_convenios.set_index("Convênio").to_dict("index")
    if not df_convenios.empty else {}
)


def construir_mapa_valores():
    mapa_valores = {}

    for convenio, dados_convenio in mapa_convenios.items():
        mapa_valores[convenio] = {}

        for codigo, dados_cbhpm in mapa_cbhpm.items():
            porte = str(dados_cbhpm.get("Porte Anest.", "")).strip()

            if porte.isdigit():
                col_an = f"AN{porte}"
                preco = limpar_moeda(dados_convenio.get(col_an, 0))
            else:
                preco = 0.0

            mapa_valores[convenio][codigo] = preco

    return mapa_valores


mapa_valores = construir_mapa_valores()

# ==========================================================
# ABAS
# ==========================================================
tab_registro, tab_dashboard = st.tabs(["📝 Novo Registro", "📊 Dashboard"])

# ==========================================================
# ABA REGISTRO
# ==========================================================
with tab_registro:

    lista_procedimentos = [
        f"{cod} - {dados.get('Procedimento','')}"
        for cod, dados in mapa_cbhpm.items()
    ]

    lista_convenios = list(mapa_convenios.keys())

    with st.form("form_cirurgia", clear_on_submit=True):

        col0, col1, col2 = st.columns(3)

        with col0:
            data_cirurgia = st.date_input("DATA", value=date.today())
        with col1:
            inicio = st.time_input("INÍCIO")
        with col2:
            termino = st.time_input("TÉRMINO")

        nome_paciente = st.text_input("NOME DO PACIENTE")
        convenio_selecionado = st.selectbox("CONVÊNIO", [""] + lista_convenios)
        procedimentos = st.multiselect("PROCEDIMENTOS", lista_procedimentos)

        submit = st.form_submit_button("💾 Salvar Registro")

    if submit:

        if not convenio_selecionado or not procedimentos:
            st.warning("Selecione convênio e procedimento.")
        else:

            duracao_str = ""

            if inicio and termino:
                t_inicio = datetime.combine(data_cirurgia, inicio)
                t_fim = datetime.combine(data_cirurgia, termino)

                if t_fim < t_inicio:
                    t_fim += pd.Timedelta(days=1)

                duracao = t_fim - t_inicio
                duracao_str = str(duracao).split(".")[0]

            nova_linha = [
                data_cirurgia.strftime("%d/%m/%Y"),
                inicio.strftime("%H:%M") if inicio else "",
                termino.strftime("%H:%M") if termino else "",
                duracao_str,
                nome_paciente,
                "\n".join(procedimentos),
                convenio_selecionado
            ]

            aba_cirurgias.append_row(
                nova_linha,
                value_input_option="USER_ENTERED"
            )

            carregar_dados.clear()
            st.success("Registro salvo com sucesso!")
            st.rerun()

# ==========================================================
# ABA DASHBOARD
# ==========================================================
with tab_dashboard:

    if df_cirurgias.empty:
        st.info("Nenhuma cirurgia registrada.")
        st.stop()

    df_cirurgias["DATA"] = pd.to_datetime(
        df_cirurgias["DATA"],
        format="%d/%m/%Y",
        errors="coerce"
    )

    data_inicio = st.date_input("De", value=df_cirurgias["DATA"].min())
    data_fim = st.date_input("Até", value=df_cirurgias["DATA"].max())

    df_filtrado = df_cirurgias[
        (df_cirurgias["DATA"] >= pd.to_datetime(data_inicio)) &
        (df_cirurgias["DATA"] <= pd.to_datetime(data_fim))
    ].copy()

    # ==============================
    # CÁLCULO DE FATURAMENTO OTIMIZADO
    # ==============================
    valores = []

    for _, row in df_filtrado.iterrows():

        convenio = str(row.get("CONVÊNIO", "")).strip()
        procs_str = str(row.get("PROCEDIMENTO", "")).strip()

        if not convenio or convenio not in mapa_valores or not procs_str:
            valores.append(0.0)
            continue

        total = 0.0
        lista_procs = procs_str.split("\n")

        for i, proc in enumerate(lista_procs):
            codigo = proc.split(" - ")[0].strip()
            preco = mapa_valores[convenio].get(codigo, 0.0)
            total += preco if i == 0 else preco * 0.5

        valores.append(total)

    df_filtrado["Valor Virtual"] = valores
    df_filtrado["Horas"] = converter_para_horas_robusto(df_filtrado["DURAÇÃO"])
    df_filtrado["R$/Hora"] = df_filtrado["Valor Virtual"] / df_filtrado["Horas"]
    df_filtrado.loc[df_filtrado["Horas"] <= 0, "R$/Hora"] = None

    # ==============================
    # MÉTRICAS
    # ==============================
    faturamento_total = df_filtrado["Valor Virtual"].sum()
    total_cirurgias = len(df_filtrado)
    ticket_medio = faturamento_total / total_cirurgias if total_cirurgias else 0
    horas_totais = df_filtrado["Horas"].sum(skipna=True)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Faturamento Total", formatar_real(faturamento_total))
    col2.metric("🏥 Nº Cirurgias", total_cirurgias)
    col3.metric("📊 Ticket Médio", formatar_real(ticket_medio))
    col4.metric("⏱ Horas Trabalhadas", f"{horas_totais:.1f} h")

    # ==============================
    # EVOLUÇÃO
    # ==============================
    st.subheader("📈 Evolução de Faturamento")
    evolucao = df_filtrado.groupby("DATA")["Valor Virtual"].sum()
    st.line_chart(evolucao)

    # ==============================
    # RESUMO POR CONVÊNIO
    # ==============================
    df_validos = df_filtrado.dropna(subset=["R$/Hora"]).copy()

    if not df_validos.empty:

        resumo = df_validos.groupby("CONVÊNIO").agg({
            "Valor Virtual": "sum",
            "Horas": "sum"
        })

        resumo["R$/Hora"] = resumo["Valor Virtual"] / resumo["Horas"]
        resumo["% Faturamento"] = (
            resumo["Valor Virtual"] / faturamento_total * 100
            if faturamento_total > 0 else 0
        )

        st.dataframe(resumo, use_container_width=True)

        def gerar_excel(df_detalhado, df_resumo):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_detalhado.to_excel(
                    writer,
                    index=False,
                    sheet_name="Dados Detalhados"
                )
                df_resumo.to_excel(
                    writer,
                    sheet_name="Resumo por Convenio"
                )
            buffer.seek(0)
            return buffer

        st.download_button(
            "📊 Baixar Relatório em Excel",
            gerar_excel(df_validos, resumo),
            file_name="relatorio_cirurgias.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )