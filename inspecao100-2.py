# Importar as Bibliotecas
import pandas as pd
import openpyxl
import plotly.express as px
import streamlit as st
import re
from datetime import datetime

# Layout da tela e Titulo
st.set_page_config(page_title="INSPEÇÃO 100% ✅", layout="wide")
st.title("INSPEÇÃO 100% DE SEGUNDA QUALIDADE 🧾") 

# CSS para customizar multiselect
st.markdown("""
    <style>
    .stMultiSelect [data-baseweb="tag"] {
        background-color: #004A99 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Competência
competencia = st.text_input("Digite a competência do arquivo? 📅", value="", placeholder="MM/YYYY")

def valida_competencia(s):
    return bool(re.match(r"^(0[1-9]|1[0-2])\/\d{4}$", s)) if s else False

# Upload
uploaded_file = st.file_uploader("Envie o arquivo Excel (.xlsx) 📂", type=["xlsx"], accept_multiple_files=False)

# Processar o arquivo
if st.button("Processar arquivo ▶️"):
    if not uploaded_file:
        st.warning("⚠️ Nenhum arquivo enviado.")
    elif not valida_competencia(competencia):
        st.warning("⚠️ Competência inválida. Use MM/YYYY.")
    else:
        try:
            # Lê o arquivo Excel
            df = pd.read_excel(uploaded_file, sheet_name="data")

            # Adiciona competência
            mes, ano = map(int, competencia.split("/"))
            data_repr = datetime(ano, mes, 1)
            df["competencia"] = pd.to_datetime(data_repr)
            df["competencia_texto"] = competencia

            # Salva na sessão
            st.session_state["master_df"] = df
            st.success(f"✅ Arquivo processado: {len(df)} linhas carregadas")

        except Exception as e:
            st.error(f"❌ Erro ao processar: {e}")

# Exibe análises se houver dados processados
if "master_df" in st.session_state:
    df = st.session_state["master_df"]

    st.markdown("---")
    st.header("📊 Análise de Dados")

    # ========== PROCESSAMENTO DOS DADOS ==========

    # DataFrame 1: OPs únicas
    df1 = df[["OP","OFICINA","LOCAL DA INSPEÇÃO","STATUS","QUANTIDADE DE PEÇAS POR OP","QUANTIDADE APRESENTADA"]]
    df1 = df1.drop_duplicates(subset=["OP","LOCAL DA INSPEÇÃO"], keep="first")
    df1["QTDE DE OP"] = 1

    # DataFrame 2: Erros de qualidade
    df2 = df[["OP","LOCAL DA INSPEÇÃO","MEDIDA","PRIMEIRA QUALIDADE","PERDAS","SEGUNDA QUALIDADE","QUANTIDADE TOTAL"]]
    df2 = df2.groupby(["LOCAL DA INSPEÇÃO","OP"], as_index=False).agg({
        "PRIMEIRA QUALIDADE": "sum",
        "PERDAS": "sum",
        "SEGUNDA QUALIDADE": "sum",
        "QUANTIDADE TOTAL": "sum"
    })

    # Renomeia coluna
    df2 = df2.rename(columns={"QUANTIDADE TOTAL": "TOTAL ERROS QUALIDADE"})

    # Merge com quantidade de peças
    df2 = df2.merge(df1[["OP", "QUANTIDADE DE PEÇAS POR OP"]], on="OP", how="left")

    # ========== FILTROS TIPO "TAGS" (MULTISELECT) ==========

    st.sidebar.header("🔍 Filtros")

    # Lista de locais e OPs
    locais_unicos = sorted(df["LOCAL DA INSPEÇÃO"].dropna().unique().tolist())
    ops_unicas = sorted(df["OP"].dropna().unique().tolist())

    # Filtros como tags (multi seleção)
    locais_selecionados = st.sidebar.multiselect(
        "Local da Inspeção (tags):",
        options=locais_unicos,
        default=locais_unicos
    )

    ops_selecionadas = st.sidebar.multiselect(
        "OP (tags):",
        options=ops_unicas,
        default=ops_unicas
    )

    # Aplica filtros no DF base (df2)
    df2_filtrado = df2.copy()

    if locais_selecionados:
        df2_filtrado = df2_filtrado[df2_filtrado["LOCAL DA INSPEÇÃO"].isin(locais_selecionados)]

    if ops_selecionadas:
        df2_filtrado = df2_filtrado[df2_filtrado["OP"].isin(ops_selecionadas)]

    # Recalcula df3 (agregado) com base no filtrado
    df3_filtrado = df2_filtrado.groupby("LOCAL DA INSPEÇÃO", as_index=False).agg({
        "TOTAL ERROS QUALIDADE": "sum",
        "QUANTIDADE DE PEÇAS POR OP": "sum"
    })

    # Calcula percentuais com base no total filtrado
    total_geral_filtrado = df3_filtrado["QUANTIDADE DE PEÇAS POR OP"].sum()

    if total_geral_filtrado > 0:
        df3_filtrado["PERCENTUAL"] = (df3_filtrado["QUANTIDADE DE PEÇAS POR OP"] / total_geral_filtrado * 100).round(2)
    else:
        df3_filtrado["PERCENTUAL"] = 0

    df3_filtrado["RÓTULO"] = (
        df3_filtrado["QUANTIDADE DE PEÇAS POR OP"].astype(int).astype(str)
        + " (" + df3_filtrado["PERCENTUAL"].astype(str) + "%)"
    )

    # ========== GRÁFICO 1: PEÇAS POR OFICINA ==========

    st.subheader("📈 Total de Peças por Oficina")

    fig1 = px.bar(
        df3_filtrado,
        x="LOCAL DA INSPEÇÃO",
        y="QUANTIDADE DE PEÇAS POR OP",
        title=f'Total de Peças por Oficina - Total Filtrado: {int(total_geral_filtrado):,} peças',
        text="RÓTULO",
        color="QUANTIDADE DE PEÇAS POR OP",
        color_continuous_scale="Blues"
    )

    fig1.update_xaxes(categoryorder='total descending')
    fig1.update_traces(textposition='outside')
    fig1.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig1, use_container_width=True, key="grafico_pecas_oficina")

    # ========== GRÁFICO 2: ERROS DE QUALIDADE ==========

    st.subheader("⚠️ Total de Erros de Qualidade por Oficina")

    total_erros_filtrado = df3_filtrado["TOTAL ERROS QUALIDADE"].sum()

    if total_erros_filtrado > 0:
        df3_filtrado["PERCENTUAL_ERROS"] = (df3_filtrado["TOTAL ERROS QUALIDADE"] / total_erros_filtrado * 100).round(2)
    else:
        df3_filtrado["PERCENTUAL_ERROS"] = 0

    df3_filtrado["RÓTULO_ERROS"] = (
        df3_filtrado["TOTAL ERROS QUALIDADE"].astype(int).astype(str)
        + " (" + df3_filtrado["PERCENTUAL_ERROS"].astype(str) + "%)"
    )

    fig2 = px.bar(
        df3_filtrado,
        x="LOCAL DA INSPEÇÃO",
        y="TOTAL ERROS QUALIDADE",
        title=f'Total de Erros de Qualidade - Total Filtrado: {int(total_erros_filtrado):,} erros',
        text="RÓTULO_ERROS",
        color="TOTAL ERROS QUALIDADE",
        color_continuous_scale="Blues"
    )

    fig2.update_xaxes(categoryorder='total descending')
    fig2.update_traces(textposition='outside')
    fig2.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig2, use_container_width=True, key="grafico_erros_oficina")

    # ========== TABELA DE DADOS ==========

    st.subheader("📋 Dados Detalhados (Filtrados)")

    st.dataframe(df2_filtrado, use_container_width=True, height=400)

    # ========== MÉTRICAS RESUMIDAS ==========

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("📦 Total de Peças", f"{int(total_geral_filtrado):,}")
    with col2:
        st.metric("⚠️ Total de Defeitos", f"{int(total_erros_filtrado):,}")
    with col3:
        st.metric("🏭 Oficinas", len(df3_filtrado))
    with col4:
        taxa_erro = (total_erros_filtrado / total_geral_filtrado * 100) if total_geral_filtrado > 0 else 0
        st.metric("📊 Taxa de Erro", f"{taxa_erro:.2f}%")

    # ========== GRÁFICO 4: PRIMEIRA QUALIDADE ==========

    df4 = df2_filtrado.groupby("LOCAL DA INSPEÇÃO", as_index=False).agg({
        "PRIMEIRA QUALIDADE": "sum"
    })

    total_primeira = df4["PRIMEIRA QUALIDADE"].sum()

    if total_primeira > 0:
        df4["PERCENTUAL"] = (df4["PRIMEIRA QUALIDADE"] / total_primeira * 100).round(2)
    else:
        df4["PERCENTUAL"] = 0

    df4["RÓTULO"] = (
        df4["PRIMEIRA QUALIDADE"].astype(int).astype(str)
        + " (" + df4["PERCENTUAL"].astype(str) + "%)"
    )

    fig4 = px.bar(
        df4,
        x="LOCAL DA INSPEÇÃO",
        y="PRIMEIRA QUALIDADE",
        title=f'Total PRIMEIRA QUALIDADE - Total Filtrado: {int(total_primeira):,}',
        text="RÓTULO",
        color="PRIMEIRA QUALIDADE",
        color_continuous_scale="Blues"
    )

    fig4.update_xaxes(categoryorder='total descending')
    fig4.update_traces(textposition='outside')
    fig4.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig4, use_container_width=True, key="grafico_primeira_qualidade")

    df4

    # ========== GRÁFICO 5: SEGUNDA QUALIDADE ==========

    df5 = df2_filtrado.groupby("LOCAL DA INSPEÇÃO", as_index=False).agg({"SEGUNDA QUALIDADE": "sum"})

    total_segunda = df5["SEGUNDA QUALIDADE"].sum()

    if total_segunda > 0:
        df5["PERCENTUAL"] = (df5["SEGUNDA QUALIDADE"] / total_segunda * 100).round(2)
    else:
        df5["PERCENTUAL"] = 0

    df5["RÓTULO"] = (
        df5["SEGUNDA QUALIDADE"].astype(int).astype(str)
        + " (" + df5["PERCENTUAL"].astype(str) + "%)"
    )

    fig5 = px.bar(
        df5,
        x="LOCAL DA INSPEÇÃO",
        y="SEGUNDA QUALIDADE",
        title=f'Total SEGUNDA QUALIDADE - Total Filtrado: {int(total_segunda):,}',
        text="RÓTULO",
        color="SEGUNDA QUALIDADE",
        color_continuous_scale="Blues"
    )

    fig5.update_xaxes(categoryorder='total descending')
    fig5.update_traces(textposition='outside')
    fig5.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig5, use_container_width=True, key="grafico_segunda_qualidade")

    df5

    # ========== GRÁFICO 6: PERDAS ==========

    df6 = df2_filtrado.groupby("LOCAL DA INSPEÇÃO", as_index=False).agg({"PERDAS": "sum"})

    total_perdas = df6["PERDAS"].sum()

    if total_perdas > 0:
        df6["PERCENTUAL"] = (df6["PERDAS"] / total_perdas * 100).round(2)
    else:
        df6["PERCENTUAL"] = 0

    df6["RÓTULO"] = (
        df6["PERDAS"].astype(int).astype(str)
        + " (" + df6["PERCENTUAL"].astype(str) + "%)"
    )

    fig6 = px.bar(
        df6,
        x="LOCAL DA INSPEÇÃO",
        y="PERDAS",
        title=f'Total PERDAS - Total Filtrado: {int(total_perdas):,}',
        text="RÓTULO",
        color="PERDAS",
        color_continuous_scale="Blues"
    )

    fig6.update_xaxes(categoryorder='total descending')
    fig6.update_traces(textposition='outside')
    fig6.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig6, use_container_width=True, key="grafico_perdas")

    df6

    # ========== PROCESSAMENTO PARA GRÁFICOS POR MEDIDA ==========

    df2m = df[["OP","LOCAL DA INSPEÇÃO","MEDIDA","PRIMEIRA QUALIDADE","PERDAS","SEGUNDA QUALIDADE","QUANTIDADE TOTAL"]]
    df2m = df2m.groupby(["LOCAL DA INSPEÇÃO","OP","MEDIDA"], as_index=False).agg({
        "PRIMEIRA QUALIDADE": "sum",
        "PERDAS": "sum",
        "SEGUNDA QUALIDADE": "sum",
        "QUANTIDADE TOTAL": "sum"
    })

    # Aplica filtros
    df2_filtradom = df2m.copy()

    if locais_selecionados:
        df2_filtradom = df2_filtradom[df2_filtradom["LOCAL DA INSPEÇÃO"].isin(locais_selecionados)]

    if ops_selecionadas:
        df2_filtradom = df2_filtradom[df2_filtradom["OP"].isin(ops_selecionadas)]

    # ========== GRÁFICO 7: PRIMEIRA QUALIDADE POR MEDIDA ==========

    df7_filtrado = df2_filtradom.groupby(["MEDIDA"], as_index=False).agg({"PRIMEIRA QUALIDADE": "sum"})

    total_geral_filtrado7 = df7_filtrado["PRIMEIRA QUALIDADE"].sum()

    if total_geral_filtrado7 > 0:
        df7_filtrado["PERCENTUAL"] = (df7_filtrado["PRIMEIRA QUALIDADE"] / total_geral_filtrado7 * 100).round(2)
    else:
        df7_filtrado["PERCENTUAL"] = 0

    df7_filtrado["RÓTULO"] = (
        df7_filtrado["PRIMEIRA QUALIDADE"].astype(int).astype(str)
        + " (" + df7_filtrado["PERCENTUAL"].astype(str) + "%)"
    )

    fig7 = px.bar(
        df7_filtrado,
        x="MEDIDA",
        y="PRIMEIRA QUALIDADE",
        title=f'Primeira Qualidade por Medida - Total: {int(total_geral_filtrado7):,}',
        text="PRIMEIRA QUALIDADE",
        color="PRIMEIRA QUALIDADE",
        color_continuous_scale="Blues"
    )

    fig7.update_xaxes(categoryorder='total descending')
    fig7.update_traces(textposition='outside')
    fig7.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig7, use_container_width=True, key="grafico_primeira_qualidade_medida")

    # ========== GRÁFICO 8: SEGUNDA QUALIDADE POR MEDIDA ==========

    df8_filtrado = df2_filtradom.groupby(["MEDIDA"], as_index=False).agg({"SEGUNDA QUALIDADE": "sum"})

    total_geral_filtrado8 = df8_filtrado["SEGUNDA QUALIDADE"].sum()

    if total_geral_filtrado8 > 0:
        df8_filtrado["PERCENTUAL"] = (df8_filtrado["SEGUNDA QUALIDADE"] / total_geral_filtrado8 * 100).round(2)
    else:
        df8_filtrado["PERCENTUAL"] = 0

    df8_filtrado["RÓTULO"] = (
        df8_filtrado["SEGUNDA QUALIDADE"].astype(int).astype(str)
        + " (" + df8_filtrado["PERCENTUAL"].astype(str) + "%)"
    )

    fig8 = px.bar(
        df8_filtrado,
        x="MEDIDA",
        y="SEGUNDA QUALIDADE",
        title=f'Segunda Qualidade por Medida - Total: {int(total_geral_filtrado8):,}',
        text="SEGUNDA QUALIDADE",
        color="SEGUNDA QUALIDADE",
        color_continuous_scale="Blues"
    )

    fig8.update_xaxes(categoryorder='total descending')
    fig8.update_traces(textposition='outside')
    fig8.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig8, use_container_width=True, key="grafico_segunda_qualidade_medida")

    # ========== GRÁFICO 9: PERDAS POR MEDIDA ==========

    df9_filtrado = df2_filtradom.groupby(["MEDIDA"], as_index=False).agg({"PERDAS": "sum"})

    total_geral_filtrado9 = df9_filtrado["PERDAS"].sum()

    if total_geral_filtrado9 > 0:
        df9_filtrado["PERCENTUAL"] = (df9_filtrado["PERDAS"] / total_geral_filtrado9 * 100).round(2)
    else:
        df9_filtrado["PERCENTUAL"] = 0

    df9_filtrado["RÓTULO"] = (
        df9_filtrado["PERDAS"].astype(int).astype(str)
        + " (" + df9_filtrado["PERCENTUAL"].astype(str) + "%)"
    )

    fig9 = px.bar(
        df9_filtrado,
        x="MEDIDA",
        y="PERDAS",
        title=f'Total de Perdas por Medida - Total: {int(total_geral_filtrado9):,}',
        text="PERDAS",
        color="PERDAS",
        color_continuous_scale="Blues"
    )

    fig9.update_xaxes(categoryorder='total descending')
    fig9.update_traces(textposition='outside')
    fig9.update_layout(height=500, showlegend=False)

    st.plotly_chart(fig9, use_container_width=True, key="grafico_perdas_medida")
