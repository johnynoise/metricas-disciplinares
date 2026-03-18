from __future__ import annotations

import argparse
import json
import math
import re
import unicodedata
from pathlib import Path
from typing import Any

import pandas as pd

TIPO_MEDIDA_MAP = {
    2: "Advertência disciplinar",
    3: "Advertência verbal",
    4: "Suspensão",
}

PESO_SEVERIDADE_MAP = {
    2: 2,
    3: 1,
    4: 4,
}

TIPO_COR_MAP = {
    "Advertência disciplinar": {"cor": "#f59e0b", "fundo": "#fff7ed", "borda": "#fdba74"},
    "Advertência verbal": {"cor": "#2563eb", "fundo": "#eff6ff", "borda": "#93c5fd"},
    "Suspensão": {"cor": "#dc2626", "fundo": "#fef2f2", "borda": "#fca5a5"},
    "Tipo não mapeado": {"cor": "#6b7280", "fundo": "#f9fafb", "borda": "#d1d5db"},
}

STOPWORDS = {
    "a", "ao", "aos", "as", "com", "como", "da", "das", "de", "do", "dos", "e", "em", "entre",
    "foi", "forma", "na", "nas", "no", "nos", "o", "os", "ou", "para", "pela", "pelas", "pelo",
    "pelos", "por", "que", "se", "sem", "ser", "sua", "seu", "suas", "seus", "um", "uma", "ao",
    "colaborador", "empregado", "empresa", "normas", "internas", "conforme", "data", "site", "deve",
    "mais", "apos", "após", "ficando", "esta", "está", "sendo", "sobre", "mesmo", "todo", "toda",
    "devido", "diante", "ocorrido", "houve", "aplicavel", "aplicável", "utilizado", "utilizacao",
    "utilização", "parte", "integrante", "instalacoes", "instalações", "disponibilizadas", "disponibilizado",
}

REGRAS_CATEGORIA = [
    (
        "Segurança operacional",
        [
            r"seguran",
            r"acident",
            r"incidente",
            r"quase acidente",
            r"mina",
            r"tr[aâ]nsito",
            r"tr[aá]fego",
            r"placa de pare",
            r"equipamento",
            r"perfura",
            r"opera(c|ç)[aã]o",
            r"munk",
            r"caminh[aã]o",
            r"risco",
        ],
    ),
    (
        "Conduta e convivência",
        [
            r"conduta",
            r"conviv",
            r"higiene",
            r"respeito",
            r"alojamento",
            r"inadequad",
            r"comportament",
        ],
    ),
    (
        "Assiduidade e jornada",
        [
            r"aus[eê]ncia",
            r"atras",
            r"falta",
            r"n[aã]o retornou",
            r"retornar",
            r"transporte",
            r"comparecimento",
            r"perda do transporte",
        ],
    ),
    (
        "Procedimento e comunicação",
        [
            r"procedimento",
            r"comunica",
            r"omitir",
            r"omiss[aã]o",
            r"plano",
            r"padr[aã]o",
            r"norma",
            r"regra",
        ],
    ),
    (
        "Patrimônio e instalações",
        [
            r"instala(c|ç)[aã]o",
            r"patrim[oô]nio",
            r"quebrou",
            r"quebrando",
            r"danific",
            r"mesa",
            r"coluna de perfura(c|ç)[aã]o",
        ],
    ),
]

HTML_CSS = """
body {
    font-family: Arial, Helvetica, sans-serif;
    margin: 0;
    padding: 24px;
    background: #f6f8fb;
    color: #1f2937;
}
.container {
    max-width: 1280px;
    margin: 0 auto;
}
.header {
    display: flex;
    justify-content: space-between;
    align-items: end;
    gap: 16px;
    margin-bottom: 24px;
}
.header h1 {
    margin: 0;
    font-size: 28px;
}
.header p {
    margin: 8px 0 0;
    color: #6b7280;
}
.section-title {
    margin: 32px 0 12px;
    font-size: 18px;
}
.cards {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
    gap: 16px;
}
.type-cards {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 16px;
}
.card {
    background: white;
    border-radius: 14px;
    padding: 18px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
}
.type-card {
    border: 1px solid #e5e7eb;
}
.type-card .label {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
}
.type-card .value {
    margin-top: 14px;
}
.type-card .meta {
    margin-top: 14px;
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 10px;
}
.type-card .meta-item {
    background: rgba(255, 255, 255, 0.62);
    border-radius: 10px;
    padding: 10px 12px;
}
.meta-item .meta-label {
    font-size: 12px;
    color: #6b7280;
}
.meta-item .meta-value {
    margin-top: 4px;
    font-size: 18px;
    font-weight: 700;
}
.badge {
    display: inline-flex;
    align-items: center;
    border-radius: 999px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 700;
}
.bar-caption {
    margin-bottom: 14px;
    color: #6b7280;
    font-size: 13px;
}
.card .label {
    color: #6b7280;
    font-size: 13px;
    margin-bottom: 8px;
}
.card .value {
    font-size: 28px;
    font-weight: bold;
}
.card .subvalue {
    margin-top: 6px;
    color: #374151;
    font-size: 13px;
}
.grid-2 {
    display: grid;
    grid-template-columns: 1.1fr 0.9fr;
    gap: 16px;
}
.panel {
    background: white;
    border-radius: 14px;
    padding: 18px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
}
.table {
    width: 100%;
    border-collapse: collapse;
    font-size: 14px;
}
.table th,
.table td {
    padding: 10px 12px;
    border-bottom: 1px solid #e5e7eb;
    text-align: left;
}
.table th {
    background: #f9fafb;
}
.bar-list {
    display: flex;
    flex-direction: column;
    gap: 10px;
}
.bar-row {
    display: grid;
    grid-template-columns: 210px 1fr 56px;
    gap: 12px;
    align-items: center;
    font-size: 14px;
}
.bar-track {
    background: #e5e7eb;
    height: 12px;
    border-radius: 999px;
    overflow: hidden;
}
.bar-fill {
    background: linear-gradient(90deg, #2563eb, #0ea5e9);
    height: 100%;
    border-radius: 999px;
}
.insights {
    margin: 0;
    padding-left: 18px;
    line-height: 1.6;
}
.footer {
    margin-top: 24px;
    color: #6b7280;
    font-size: 12px;
}
@media (max-width: 900px) {
    .grid-2 {
        grid-template-columns: 1fr;
    }
    .bar-row {
        grid-template-columns: 1fr;
    }
}
"""


def normalizar_texto(valor: Any) -> str:
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return ""
    texto = str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return texto.lower()


def classificar_categoria(anotacao: Any) -> str:
    texto = normalizar_texto(anotacao)
    if not texto:
        return "Não informado"

    melhor_categoria = "Outros"
    melhor_score = 0
    for categoria, padroes in REGRAS_CATEGORIA:
        score = sum(1 for padrao in padroes if re.search(padrao, texto))
        if score > melhor_score:
            melhor_score = score
            melhor_categoria = categoria
    return melhor_categoria


def extrair_palavras_frequentes(series: pd.Series, limite: int = 15) -> list[dict[str, Any]]:
    frequencias: dict[str, int] = {}
    for texto in series.fillna(""):
        bruto = normalizar_texto(texto)
        palavras = re.findall(r"[a-zA-Zà-úÀ-Ú]{4,}", bruto)
        for palavra in palavras:
            if palavra in STOPWORDS or palavra.isdigit():
                continue
            frequencias[palavra] = frequencias.get(palavra, 0) + 1

    ranking = sorted(frequencias.items(), key=lambda item: (-item[1], item[0]))[:limite]
    return [{"palavra": palavra, "frequencia": freq} for palavra, freq in ranking]


def formatar_percentual(valor: float) -> str:
    return f"{valor * 100:.1f}%"


def formatar_data(valor: pd.Timestamp | None) -> str:
    if valor is None or pd.isna(valor):
        return "-"
    return valor.strftime("%d/%m/%Y")


def obter_estilo_tipo(tipo: str) -> dict[str, str]:
    return TIPO_COR_MAP.get(tipo, TIPO_COR_MAP["Tipo não mapeado"])


def montar_insights(
    resumo: dict[str, Any],
    por_tipo: pd.DataFrame,
    por_mes: pd.DataFrame,
    por_categoria: pd.DataFrame,
) -> list[str]:
    insights: list[str] = []
    if not por_tipo.empty:
        principal_tipo = por_tipo.iloc[0]
        insights.append(
            f"{principal_tipo['TipoDescricao']} representa {principal_tipo['Percentual']} das medidas, sinalizando o perfil predominante de atuação disciplinar."
        )
    if not por_mes.empty:
        pico = por_mes.sort_values("Quantidade", ascending=False).iloc[0]
        insights.append(
            f"O mês com maior volume foi {pico['AnoMesReferencia']} com {int(pico['Quantidade'])} registro(s)."
        )
    if resumo["taxa_reincidencia_medidas"] > 0:
        insights.append(
            f"Há {resumo['reincidencias']} medida(s) com sequência maior que 1, equivalente a {resumo['taxa_reincidencia_medidas_percentual']} do total."
        )
    else:
        insights.append("Não há medidas com sequência maior que 1 na base atual, o que sugere baixa repetição formal de casos.")
    if not por_categoria.empty:
        principal_categoria = por_categoria.iloc[0]
        insights.append(
            f"A categoria mais frequente é {principal_categoria['CategoriaMotivo']}, com {int(principal_categoria['Quantidade'])} ocorrência(s), indicando foco prioritário de prevenção."
        )
    insights.append(
        "Recomenda-se acompanhar mensalmente o mix por severidade, sequência das medidas e categorias de motivo para agir antes da escalada disciplinar."
    )
    return insights


def criar_barras_html(df: pd.DataFrame, label_col: str, value_col: str, color_map: dict[str, dict[str, str]] | None = None) -> str:
    if df.empty:
        return "<p>Sem dados suficientes.</p>"
    maior = max(df[value_col].max(), 1)
    linhas = []
    for _, row in df.iterrows():
        largura = (float(row[value_col]) / float(maior)) * 100
        estilo = (color_map or {}).get(str(row[label_col]), {})
        cor = estilo.get("cor", "#2563eb")
        linhas.append(
            f"<div class='bar-row'><div>{row[label_col]}</div><div class='bar-track'><div class='bar-fill' style='width:{largura:.1f}%; background:{cor};'></div></div><div>{int(row[value_col])}</div></div>"
        )
    return "<div class='bar-list'>" + "".join(linhas) + "</div>"


def criar_tabela_html(df: pd.DataFrame) -> str:
    if df.empty:
        return "<p>Sem dados suficientes.</p>"
    colunas = list(df.columns)
    thead = "".join(f"<th>{coluna}</th>" for coluna in colunas)
    linhas = []
    for _, row in df.iterrows():
        tds = "".join(f"<td>{row[coluna]}</td>" for coluna in colunas)
        linhas.append(f"<tr>{tds}</tr>")
    return f"<table class='table'><thead><tr>{thead}</tr></thead><tbody>{''.join(linhas)}</tbody></table>"


def criar_cards_tipo_html(por_tipo: pd.DataFrame, total_medidas: int) -> str:
    if por_tipo.empty:
        return "<p>Sem dados suficientes.</p>"

    blocos = []
    for _, row in por_tipo.iterrows():
        tipo = str(row["TipoDescricao"])
        quantidade = int(row["Quantidade"])
        percentual = row["Percentual"]
        estilo = obter_estilo_tipo(tipo)
        incidencia = "Alta concentração" if total_medidas and quantidade / total_medidas >= 0.5 else "Monitorar"
        blocos.append(
            f"""
            <div class='card type-card' style='background:{estilo['fundo']}; border-color:{estilo['borda']}'>
                <div class='label'>
                    <span>{tipo}</span>
                    <span class='badge' style='background:{estilo['cor']}; color:white'>{percentual}</span>
                </div>
                <div class='value' style='color:{estilo['cor']}'>{quantidade}</div>
                <div class='subvalue'>medida(s) registradas no período</div>
                <div class='meta'>
                    <div class='meta-item'>
                        <div class='meta-label'>Participação</div>
                        <div class='meta-value'>{percentual}</div>
                    </div>
                    <div class='meta-item'>
                        <div class='meta-label'>Leitura RH</div>
                        <div class='meta-value'>{incidencia}</div>
                    </div>
                </div>
            </div>
            """
        )
    return "<div class='type-cards'>" + "".join(blocos) + "</div>"


def gerar_dashboard_html(
    arquivo_saida: Path,
    resumo: dict[str, Any],
    insights: list[str],
    por_tipo: pd.DataFrame,
    por_mes: pd.DataFrame,
    por_categoria: pd.DataFrame,
    top_colaboradores: pd.DataFrame,
    palavras: list[dict[str, Any]],
) -> None:
    palavras_df = pd.DataFrame(palavras)
    cards = [
        ("Total de medidas", resumo["total_medidas"], f"{resumo['colaboradores_unicos']} colaborador(es) impactado(s)"),
        ("Período analisado", resumo["periodo_texto"], f"{resumo['meses_com_registro']} mês(es) com registro"),
        ("Sequência > 1", resumo["reincidencias"], f"Taxa sobre medidas: {resumo['taxa_reincidencia_medidas_percentual']}"),
        ("Severidade média", resumo["severidade_media"], f"Índice ponderado de {resumo['indice_severidade_ponderado']:.2f}"),
        ("Tipo dominante", resumo["tipo_mais_frequente"], f"{resumo['tipo_mais_frequente_percentual']} das medidas"),
        ("Categoria dominante", resumo["categoria_mais_frequente"], f"{resumo['categoria_mais_frequente_quantidade']} ocorrência(s)"),
    ]
    cards_html = "".join(
        f"<div class='card'><div class='label'>{label}</div><div class='value'>{value}</div><div class='subvalue'>{sub}</div></div>"
        for label, value, sub in cards
    )
    cards_tipo_html = criar_cards_tipo_html(por_tipo, resumo["total_medidas"])
    insights_html = "".join(f"<li>{insight}</li>" for insight in insights)

    html = f"""
<!DOCTYPE html>
<html lang='pt-BR'>
<head>
    <meta charset='utf-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1'>
    <title>KPIs de Medidas Disciplinares</title>
    <style>{HTML_CSS}</style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <div>
                <h1>KPIs de Medidas Disciplinares</h1>
                <p>{resumo['empresa_principal']} · Atualizado em {resumo['data_geracao']}</p>
            </div>
            <div>
                <p>Base: {resumo['arquivo_origem']}</p>
            </div>
        </div>

        <div class='cards'>{cards_html}</div>

        <h2 class='section-title'>Painel por tipo de medida</h2>
        <div class='panel'>
            <div class='bar-caption'>Blocos executivos para destacar o peso de cada tipo disciplinar na base atual.</div>
            {cards_tipo_html}
        </div>

        <h2 class='section-title'>Insights executivos</h2>
        <div class='panel'>
            <ul class='insights'>{insights_html}</ul>
        </div>

        <div class='grid-2'>
            <div>
                <h2 class='section-title'>Distribuição por tipo</h2>
                <div class='panel'>
                    {criar_barras_html(por_tipo, 'TipoDescricao', 'Quantidade', TIPO_COR_MAP)}
                </div>
            </div>
            <div>
                <h2 class='section-title'>Distribuição por categoria</h2>
                <div class='panel'>
                    {criar_barras_html(por_categoria, 'CategoriaMotivo', 'Quantidade')}
                </div>
            </div>
        </div>

        <div class='grid-2'>
            <div>
                <h2 class='section-title'>Evolução mensal</h2>
                <div class='panel'>{criar_tabela_html(por_mes)}</div>
            </div>
            <div>
                <h2 class='section-title'>Colaboradores com maior severidade</h2>
                <div class='panel'>{criar_tabela_html(top_colaboradores)}</div>
            </div>
        </div>

        <div class='grid-2'>
            <div>
                <h2 class='section-title'>Palavras recorrentes nas anotações</h2>
                <div class='panel'>{criar_tabela_html(palavras_df)}</div>
            </div>
            <div>
                <h2 class='section-title'>Detalhe por tipo</h2>
                <div class='panel'>{criar_tabela_html(por_tipo)}</div>
            </div>
        </div>

        <div class='footer'>
            Assunção padrão do RH: 002 = Advertência disciplinar, 003 = Advertência verbal e 004 = Suspensão. A coluna Sequencia representa a ordem da medida aplicada ao colaborador; valores maiores que 1 indicam repetição.
        </div>
    </div>
</body>
</html>
"""
    arquivo_saida.write_text(html, encoding="utf-8")


def preparar_dataframe(caminho_excel: Path) -> pd.DataFrame:
    df = pd.read_excel(caminho_excel)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
    df.columns = [str(coluna).strip() for coluna in df.columns]

    obrigatorias = {"Empresa", "Data", "Funcionario", "Sequencia", "Tipo", "Matricula", "Anotacoes"}
    faltantes = obrigatorias.difference(df.columns)
    if faltantes:
        raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(sorted(faltantes))}")

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    if df["Data"].isna().all():
        df["Data"] = pd.to_datetime(df["Data"], unit="D", origin="1899-12-30", errors="coerce")

    df["Tipo"] = pd.to_numeric(df["Tipo"], errors="coerce").astype("Int64")
    df["Sequencia"] = pd.to_numeric(df["Sequencia"], errors="coerce").astype("Int64")
    df["Matricula"] = pd.to_numeric(df["Matricula"], errors="coerce").astype("Int64")

    df["TipoDescricao"] = df["Tipo"].map(TIPO_MEDIDA_MAP).fillna("Tipo não mapeado")
    df["PesoSeveridade"] = df["Tipo"].map(PESO_SEVERIDADE_MAP).fillna(0)
    df["CategoriaMotivo"] = df["Anotacoes"].apply(classificar_categoria)
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["EhReincidencia"] = df["Sequencia"].fillna(1).gt(1)
    df["DataTexto"] = df["Data"].dt.strftime("%d/%m/%Y")
    return df.sort_values(["Data", "Tipo", "Funcionario"], ascending=[True, True, True]).reset_index(drop=True)


def calcular_kpis(df: pd.DataFrame, caminho_excel: Path) -> tuple[dict[str, Any], dict[str, pd.DataFrame], list[dict[str, Any]], list[str]]:
    total_medidas = int(len(df))
    colaboradores_unicos = int(df["Matricula"].dropna().nunique()) if "Matricula" in df else int(df["Funcionario"].nunique())
    reincidencias = int(df["EhReincidencia"].sum())
    taxa_reincidencia = (reincidencias / total_medidas) if total_medidas else 0.0
    severidade_media = float(df["PesoSeveridade"].mean()) if total_medidas else 0.0

    por_tipo = (
        df.groupby("TipoDescricao", dropna=False)
        .size()
        .reset_index(name="Quantidade")
        .sort_values("Quantidade", ascending=False)
        .reset_index(drop=True)
    )
    por_tipo["Percentual"] = (por_tipo["Quantidade"] / total_medidas).map(formatar_percentual)

    por_mes = (
        df.groupby("AnoMes", dropna=False)
        .size()
        .reset_index(name="Quantidade")
        .sort_values("AnoMes")
        .reset_index(drop=True)
    )
    por_mes["AnoMesReferencia"] = pd.to_datetime(por_mes["AnoMes"] + "-01").dt.strftime("%m/%Y")
    por_mes = por_mes[["AnoMesReferencia", "Quantidade"]]

    por_categoria = (
        df.groupby("CategoriaMotivo", dropna=False)
        .size()
        .reset_index(name="Quantidade")
        .sort_values("Quantidade", ascending=False)
        .reset_index(drop=True)
    )
    por_categoria["Percentual"] = (por_categoria["Quantidade"] / total_medidas).map(formatar_percentual)

    top_colaboradores = (
        df.groupby(["Funcionario", "Matricula"], dropna=False)
        .agg(
            Quantidade=("Funcionario", "size"),
            SeveridadeMaxima=("PesoSeveridade", "max"),
            UltimaOcorrencia=("Data", "max"),
            Reincidente=("EhReincidencia", "max"),
        )
        .reset_index()
        .sort_values(["SeveridadeMaxima", "Quantidade", "UltimaOcorrencia"], ascending=[False, False, False])
        .head(10)
        .reset_index(drop=True)
    )
    top_colaboradores["UltimaOcorrencia"] = top_colaboradores["UltimaOcorrencia"].dt.strftime("%d/%m/%Y")
    top_colaboradores["Reincidente"] = top_colaboradores["Reincidente"].map({True: "Sim", False: "Não"})

    tipo_mais_frequente = por_tipo.iloc[0] if not por_tipo.empty else None
    categoria_mais_frequente = por_categoria.iloc[0] if not por_categoria.empty else None

    empresa_principal = df["Empresa"].mode().iloc[0] if not df["Empresa"].mode().empty else "Empresa não identificada"
    periodo_inicio = df["Data"].min()
    periodo_fim = df["Data"].max()

    resumo = {
        "arquivo_origem": caminho_excel.name,
        "empresa_principal": empresa_principal,
        "data_geracao": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"),
        "total_medidas": total_medidas,
        "colaboradores_unicos": colaboradores_unicos,
        "reincidencias": reincidencias,
        "taxa_reincidencia_medidas": taxa_reincidencia,
        "taxa_reincidencia_medidas_percentual": formatar_percentual(taxa_reincidencia),
        "indice_severidade_ponderado": severidade_media,
        "severidade_media": f"{severidade_media:.2f}",
        "tipo_mais_frequente": tipo_mais_frequente["TipoDescricao"] if tipo_mais_frequente is not None else "-",
        "tipo_mais_frequente_percentual": tipo_mais_frequente["Percentual"] if tipo_mais_frequente is not None else "-",
        "categoria_mais_frequente": categoria_mais_frequente["CategoriaMotivo"] if categoria_mais_frequente is not None else "-",
        "categoria_mais_frequente_quantidade": int(categoria_mais_frequente["Quantidade"]) if categoria_mais_frequente is not None else 0,
        "periodo_inicio": formatar_data(periodo_inicio),
        "periodo_fim": formatar_data(periodo_fim),
        "periodo_texto": f"{formatar_data(periodo_inicio)} a {formatar_data(periodo_fim)}",
        "meses_com_registro": int(df["AnoMes"].nunique()),
    }

    palavras = extrair_palavras_frequentes(df["Anotacoes"])
    insights = montar_insights(resumo, por_tipo, por_mes, por_categoria)

    tabelas = {
        "base_tratada": df,
        "por_tipo": por_tipo,
        "por_mes": por_mes,
        "por_categoria": por_categoria,
        "top_colaboradores": top_colaboradores,
        "palavras_frequentes": pd.DataFrame(palavras),
    }
    return resumo, tabelas, palavras, insights


def exportar_resultados(
    output_dir: Path,
    resumo: dict[str, Any],
    tabelas: dict[str, pd.DataFrame],
    palavras: list[dict[str, Any]],
    insights: list[str],
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    tabelas["base_tratada"].to_csv(output_dir / "medidas_tratadas.csv", index=False, encoding="utf-8-sig")
    with (output_dir / "kpis_resumo.json").open("w", encoding="utf-8") as arquivo:
        json.dump({"resumo": resumo, "insights": insights, "palavras_frequentes": palavras}, arquivo, ensure_ascii=False, indent=2)

    with pd.ExcelWriter(output_dir / "kpis_medidas_disciplinares.xlsx", engine="openpyxl") as writer:
        pd.DataFrame([resumo]).to_excel(writer, sheet_name="resumo", index=False)
        for nome, tabela in tabelas.items():
            tabela.to_excel(writer, sheet_name=nome[:31], index=False)

    gerar_dashboard_html(
        output_dir / "dashboard_kpis_medidas_disciplinares.html",
        resumo,
        insights,
        tabelas["por_tipo"],
        tabelas["por_mes"],
        tabelas["por_categoria"],
        tabelas["top_colaboradores"],
        palavras,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera KPIs de RH a partir de planilha de medidas disciplinares.")
    parser.add_argument(
        "--input",
        default="medidas_disciplinares_limpo.xlsx",
        help="Caminho do arquivo Excel de entrada.",
    )
    parser.add_argument(
        "--output-dir",
        default="saida_kpis_rh",
        help="Pasta onde os arquivos gerados serão gravados.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    caminho_excel = Path(args.input).resolve()
    output_dir = Path(args.output_dir).resolve()

    if not caminho_excel.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_excel}")

    df = preparar_dataframe(caminho_excel)
    resumo, tabelas, palavras, insights = calcular_kpis(df, caminho_excel)
    exportar_resultados(output_dir, resumo, tabelas, palavras, insights)

    print("KPIs gerados com sucesso.")
    print(f"Entrada: {caminho_excel}")
    print(f"Saída:   {output_dir}")
    print(f"Total de medidas: {resumo['total_medidas']}")
    print(f"Sequência > 1: {resumo['reincidencias']} ({resumo['taxa_reincidencia_medidas_percentual']})")
    print(f"Tipo mais frequente: {resumo['tipo_mais_frequente']} ({resumo['tipo_mais_frequente_percentual']})")
    print(f"Categoria mais frequente: {resumo['categoria_mais_frequente']}")


if __name__ == "__main__":
    main()
