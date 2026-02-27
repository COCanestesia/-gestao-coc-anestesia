import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread
from io import BytesIO

# PDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4

# Excel
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Gestão COC Anestesia", layout="wide")
st.title("Registro de Cirurgias e Dashboard Financeiro")

try:
    client = ()

    URL_PLANILHA = "COLE_AQUI_A_URL_REAL_DA_PLANILHA"

    planilha = client.open_by_url(URL_PLANILHA)

    aba_cirurgias = planilha.worksheet("CIRURGIAS")
    aba_convenios = planilha.worksheet("Página2")
    aba_cbhpm = planilha.worksheet("Página3")

    df_cirurgias = pd.DataFrame(aba_cirurgias.get_all_records())
    df_convenios = pd.DataFrame(aba_convenios.get_all_records())
    df_cbhpm = pd.DataFrame(aba_cbhpm.get_all_records())

except Exception as e:
    st.error(f"Erro de conexão com as planilhas: {e}")
    st.stop()

# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================
def limpar_moeda(valor):
    if pd.isna(valor) or str(valor).strip() in ['-', '', '0']:
        return 0.0
    valor_str = str(valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
    try:
        return float(valor_str)
    except:
        return 0.0

def formatar_real(valor):
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def converter_para_horas(duracao_str):
    try:
        texto = str(duracao_str).strip()
        if not texto or texto == "nan":
            return None
        partes = texto.split(':')
        return int(partes[0]) + (int(partes[1]) / 60.0)
    except:
        return None

mapa_cbhpm = df_cbhpm.set_index('Código').to_dict('index')
mapa_convenios = df_convenios.set_index('Convênio').to_dict('index')

def calcular_faturamento_memoria(row):
    convenio = str(row.get('CONVÊNIO', '')).strip()
    procs_str = str(row.get('PROCEDIMENTO', '')).strip()

    if not convenio or not procs_str or procs_str == 'nan':
        return 0.0

    linha_convenio = mapa_convenios.get(convenio)
    if not linha_convenio:
        return 0.0

    valor_total = 0.0
    lista_procs = procs_str.split('\n')

    for i, proc in enumerate(lista_procs):
        codigo = proc.split(" - ")[0].strip()
        linha_cbhpm = mapa_cbhpm.get(codigo)
        if not linha_cbhpm:
            continue

        porte = str(linha_cbhpm.get('Porte Anest.', '')).strip()
        preco = 0.0

        if porte.isdigit():
            col_an = f"AN{porte}"
            if col_an in linha_convenio:
                preco = limpar_moeda(linha_convenio[col_an])

        if i == 0:
            valor_total += preco
        else:
            valor_total += preco * 0.5

    return valor_total

# ==========================================================
# DASHBOARD
# ==========================================================
st.subheader("🏆 Ranking de Rentabilidade")

if df_cirurgias.empty:
    st.info("Nenhuma cirurgia registrada.")
else:

    df_cirurgias['Valor Virtual'] = df_cirurgias.apply(calcular_faturamento_memoria, axis=1)

    if 'DURAÇÃO' not in df_cirurgias.columns:
        st.error("Coluna DURAÇÃO não encontrada.")
        st.stop()

    df_cirurgias['Horas'] = df_cirurgias['DURAÇÃO'].apply(converter_para_horas)

    df_cirurgias['R$/Hora'] = df_cirurgias.apply(
        lambda row: row['Valor Virtual'] / row['Horas']
        if row['Horas'] and row['Horas'] > 0 else None,
        axis=1
    )

    faturamento_total = df_cirurgias['Valor Virtual'].sum()
    total_cirurgias = len(df_cirurgias)
    ticket_medio = faturamento_total / total_cirurgias if total_cirurgias else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("💰 Faturamento Total", formatar_real(faturamento_total))
    col2.metric("🏥 Nº Cirurgias", total_cirurgias)
    col3.metric("📊 Ticket Médio", formatar_real(ticket_medio))

    df_validos = df_cirurgias.dropna(subset=['R$/Hora']).copy()

    if not df_validos.empty:

        st.subheader("📊 Ranking por R$/Hora")
        df_ranking = df_validos.sort_values(by='R$/Hora', ascending=False)
        st.dataframe(df_ranking, use_container_width=True)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_ranking.to_excel(writer, index=False, sheet_name='Ranking')

        buffer.seek(0)

        st.download_button(
            "⬇️ Baixar Relatório Excel",
            data=buffer,
            file_name="Relatorio_COC_Anestesia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ==========================================================
        # ANÁLISE POR CONVÊNIO
        # ==========================================================
        st.subheader("🏥 Análise Estratégica por Convênio")

        resumo_convenio = (
            df_validos
            .groupby('CONVÊNIO')
            .agg({'Valor Virtual': 'sum', 'Horas': 'sum'})
        )

        resumo_convenio['R$/Hora'] = resumo_convenio['Valor Virtual'] / resumo_convenio['Horas']
        resumo_convenio['% Faturamento'] = resumo_convenio['Valor Virtual'] / faturamento_total * 100
        resumo_convenio = resumo_convenio.sort_values(by='Valor Virtual', ascending=False)

        convenio_top = resumo_convenio.index[0]
        valor_top = resumo_convenio.iloc[0]['Valor Virtual']

        convenio_rentavel = resumo_convenio.sort_values(by='R$/Hora', ascending=False).index[0]
        valor_rentavel = resumo_convenio.sort_values(by='R$/Hora', ascending=False).iloc[0]['R$/Hora']

        convenio_menos = resumo_convenio.sort_values(by='R$/Hora').index[0]
        valor_menos = resumo_convenio.sort_values(by='R$/Hora').iloc[0]['R$/Hora']

        c1, c2, c3 = st.columns(3)
        c1.metric("🥇 Maior Faturamento", convenio_top, formatar_real(valor_top))
        c2.metric("⚡ Mais Rentável (R$/Hora)", convenio_rentavel, formatar_real(valor_rentavel))
        c3.metric("📉 Menos Rentável (R$/Hora)", convenio_menos, formatar_real(valor_menos))

        st.subheader("📊 Faturamento por Convênio")
        st.bar_chart(resumo_convenio['Valor Virtual'])

        st.subheader("⚡ Rentabilidade por Hora")
        st.bar_chart(resumo_convenio['R$/Hora'])

        resumo_exibicao = resumo_convenio.copy()
        resumo_exibicao['Valor Virtual'] = resumo_exibicao['Valor Virtual'].apply(formatar_real)
        resumo_exibicao['R$/Hora'] = resumo_exibicao['R$/Hora'].apply(formatar_real)
        resumo_exibicao['% Faturamento'] = resumo_exibicao['% Faturamento'].apply(lambda x: f"{x:.1f}%")

        st.subheader("📋 Tabela Estratégica Completa")
        st.dataframe(resumo_exibicao, use_container_width=True)