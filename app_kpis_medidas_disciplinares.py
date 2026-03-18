from __future__ import annotations

import io
import json
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from gerar_kpis_medidas_disciplinares import (
    TIPO_COR_MAP,
    calcular_kpis,
    exportar_resultados,
    preparar_dataframe,
)

st.set_page_config(
    page_title="Painel RH · Medidas Disciplinares",
    page_icon="📊",
    layout="wide",
)


def _inject_css() -> None:
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 1.4rem;
                padding-bottom: 2rem;
                max-width: 1400px;
            }
            .hero-card {
                background: linear-gradient(135deg, #0f172a 0%, #1e293b 35%, #2563eb 100%);
                color: #fff;
                border-radius: 18px;
                padding: 22px 24px;
                margin-bottom: 14px;
                box-shadow: 0 12px 24px rgba(15, 23, 42, 0.22);
            }
            .hero-card h1 {
                margin: 0;
                font-size: 28px;
            }
            .hero-card p {
                margin: 8px 0 0;
                color: #cbd5e1;
                font-size: 14px;
            }
            div[data-testid="metric-container"] {
                border: 1px solid #e2e8f0;
                border-radius: 14px;
                padding: 12px 14px;
                background: #ffffff;
                box-shadow: 0 3px 14px rgba(15, 23, 42, 0.06);
            }
            .section-title {
                margin-top: 0.6rem;
                margin-bottom: 0.2rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _salvar_upload_temporario(upload: st.runtime.uploaded_file_manager.UploadedFile) -> Path:
    sufixo = Path(upload.name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=sufixo) as tmp:
        tmp.write(upload.getbuffer())
        return Path(tmp.name)


def _filtrar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    filtro_empresa = st.sidebar.multiselect(
        "Empresa",
        options=sorted(df["Empresa"].dropna().astype(str).unique().tolist()),
        default=sorted(df["Empresa"].dropna().astype(str).unique().tolist()),
    )

    filtro_tipo = st.sidebar.multiselect(
        "Tipo de medida",
        options=sorted(df["TipoDescricao"].dropna().astype(str).unique().tolist()),
        default=sorted(df["TipoDescricao"].dropna().astype(str).unique().tolist()),
    )

    dt_min = df["Data"].min().date() if not df["Data"].isna().all() else None
    dt_max = df["Data"].max().date() if not df["Data"].isna().all() else None

    intervalo = None
    if dt_min and dt_max:
        intervalo = st.sidebar.date_input(
            "Período",
            value=(dt_min, dt_max),
            min_value=dt_min,
            max_value=dt_max,
        )

    filtrado = df.copy()
    if filtro_empresa:
        filtrado = filtrado[filtrado["Empresa"].astype(str).isin(filtro_empresa)]
    if filtro_tipo:
        filtrado = filtrado[filtrado["TipoDescricao"].astype(str).isin(filtro_tipo)]

    if intervalo and len(intervalo) == 2:
        inicio, fim = pd.to_datetime(intervalo[0]), pd.to_datetime(intervalo[1])
        filtrado = filtrado[(filtrado["Data"] >= inicio) & (filtrado["Data"] <= fim)]

    return filtrado.reset_index(drop=True)


def _fig_tipo(por_tipo: pd.DataFrame):
    if por_tipo.empty:
        return None

    cores = {
        tipo: TIPO_COR_MAP.get(tipo, TIPO_COR_MAP["Tipo não mapeado"])["cor"]
        for tipo in por_tipo["TipoDescricao"].tolist()
    }

    fig = px.bar(
        por_tipo,
        x="TipoDescricao",
        y="Quantidade",
        text="Percentual",
        color="TipoDescricao",
        color_discrete_map=cores,
    )
    fig.update_layout(
        showlegend=False,
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis_title="",
        yaxis_title="Quantidade",
    )
    fig.update_traces(textposition="outside")
    return fig


def _fig_categoria(por_categoria: pd.DataFrame):
    if por_categoria.empty:
        return None

    fig = px.pie(
        por_categoria,
        names="CategoriaMotivo",
        values="Quantidade",
        hole=0.62,
    )
    fig.update_layout(margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return fig


def _fig_mes(por_mes: pd.DataFrame):
    if por_mes.empty:
        return None

    base = por_mes.copy()
    base["RefDate"] = pd.to_datetime(base["AnoMesReferencia"], format="%m/%Y", errors="coerce")
    base = base.sort_values("RefDate")

    fig = px.line(
        base,
        x="AnoMesReferencia",
        y="Quantidade",
        markers=True,
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis_title="Mês",
        yaxis_title="Medidas",
    )
    return fig


def _gerar_zip_resultados(
    resumo: dict,
    tabelas: dict[str, pd.DataFrame],
    palavras: list[dict],
    insights: list[str],
) -> bytes:
    with tempfile.TemporaryDirectory() as tmp_dir:
        pasta = Path(tmp_dir)
        exportar_resultados(pasta, resumo, tabelas, palavras, insights)

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for arquivo in pasta.iterdir():
                zf.write(arquivo, arcname=arquivo.name)
        buffer.seek(0)
        return buffer.getvalue()


def _render_kpis(resumo: dict) -> None:
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total medidas", resumo["total_medidas"])
    c2.metric("Colaboradores", resumo["colaboradores_unicos"])
    c3.metric("Reincidências", resumo["reincidencias"], resumo["taxa_reincidencia_medidas_percentual"])
    c4.metric("Severidade média", resumo["severidade_media"])
    c5.metric("Tipo dominante", resumo["tipo_mais_frequente"], resumo["tipo_mais_frequente_percentual"])
    c6.metric("Categoria dominante", resumo["categoria_mais_frequente"])


def _render_resultados(
    resumo: dict,
    tabelas: dict[str, pd.DataFrame],
    palavras: list[dict],
    insights: list[str],
) -> None:
    _render_kpis(resumo)

    col_a, col_b = st.columns((1.1, 0.9))
    with col_a:
        st.markdown("### Distribuição por tipo")
        fig_tipo = _fig_tipo(tabelas["por_tipo"])
        if fig_tipo is None:
            st.info("Sem dados para o gráfico de tipo.")
        else:
            st.plotly_chart(fig_tipo, use_container_width=True)

    with col_b:
        st.markdown("### Distribuição por categoria")
        fig_cat = _fig_categoria(tabelas["por_categoria"])
        if fig_cat is None:
            st.info("Sem dados para o gráfico de categorias.")
        else:
            st.plotly_chart(fig_cat, use_container_width=True)

    col_c, col_d = st.columns((1.2, 0.8))
    with col_c:
        st.markdown("### Evolução mensal")
        fig_mes = _fig_mes(tabelas["por_mes"])
        if fig_mes is None:
            st.info("Sem dados para evolução mensal.")
        else:
            st.plotly_chart(fig_mes, use_container_width=True)

    with col_d:
        st.markdown("### Insights executivos")
        for item in insights:
            st.write(f"• {item}")

    st.markdown("### Top colaboradores por severidade")
    st.dataframe(tabelas["top_colaboradores"], use_container_width=True, hide_index=True)

    st.markdown("### Palavras recorrentes")
    st.dataframe(pd.DataFrame(palavras), use_container_width=True, hide_index=True)


def main() -> None:
    _inject_css()

    st.markdown(
        """
        <div class="hero-card">
            <h1>Painel de KPIs · Medidas Disciplinares</h1>
            <p>Importe o Excel do RH, aplique filtros e gere resultados executivos em CSV, JSON, XLSX e HTML.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.header("Importação e filtros")

    upload = st.sidebar.file_uploader(
        "Arquivo Excel de medidas disciplinares",
        type=["xlsx", "xls"],
        help="Campos esperados: Empresa, Data, Funcionario, Sequencia, Tipo, Matricula, Anotacoes.",
    )

    if not upload:
        st.info("Envie um arquivo Excel para iniciar a análise.")
        return

    caminho_arquivo = _salvar_upload_temporario(upload)

    try:
        df = preparar_dataframe(caminho_arquivo)
    except Exception as exc:
        st.error(f"Falha ao processar o arquivo: {exc}")
        return

    st.sidebar.success(f"Arquivo carregado: {upload.name}")
    df_filtrado = _filtrar_dataframe(df)

    if df_filtrado.empty:
        st.warning("Os filtros retornaram base vazia. Ajuste os filtros na barra lateral.")
        return

    resumo, tabelas, palavras, insights = calcular_kpis(df_filtrado, Path(upload.name))

    st.markdown("### Visão geral")
    st.caption(
        f"Período: {resumo['periodo_texto']} · Empresa principal: {resumo['empresa_principal']} · "
        f"Registros no recorte: {resumo['total_medidas']}"
    )

    _render_resultados(resumo, tabelas, palavras, insights)

    st.markdown("### Exportação")
    zip_bytes = _gerar_zip_resultados(resumo, tabelas, palavras, insights)
    st.download_button(
        "Baixar pacote de resultados (.zip)",
        data=zip_bytes,
        file_name="kpis_medidas_disciplinares.zip",
        mime="application/zip",
        use_container_width=True,
    )

    pasta_destino = st.text_input(
        "Opcional: salvar também em pasta local",
        value=str(Path.cwd() / "saida_kpis_rh"),
    )
    if st.button("Salvar resultados na pasta local", use_container_width=True):
        try:
            destino = Path(pasta_destino).expanduser().resolve()
            exportar_resultados(destino, resumo, tabelas, palavras, insights)
            st.success(f"Resultados salvos em: {destino}")
        except Exception as exc:
            st.error(f"Não foi possível salvar os arquivos localmente: {exc}")

    with st.expander("Visualizar resumo JSON"):
        st.code(json.dumps({"resumo": resumo, "insights": insights}, ensure_ascii=False, indent=2), language="json")


if __name__ == "__main__":
    main()
