import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="Meu Controle Financeiro",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------- Paleta de cores ----------
AZUL_LILAS = "#A5BEFA"
VINHO = "#451531"
CIANO = "#64B7CC"
PINK = "#FF3877"
ALERTA = "#B3093F"

TITULO_COR = VINHO
PRIMARIO = CIANO
FUNDO_GRAFICO = "#f3f3f3"

PALETA_CATEGORIAS = [AZUL_LILAS, VINHO, CIANO, PINK]
PALETA_PAGAMENTO = [AZUL_LILAS, CIANO, PINK]
PALETA_METRICAS = [CIANO, AZUL_LILAS, PINK, VINHO]
PALETA_CONTAS = [CIANO, PINK, AZUL_LILAS, VINHO]

# ---------- CSS ----------
st.markdown(f"""
<style>
    .stApp {{ background: #F7F8FA; }}
    .block-container {{ max-width: 1450px; padding-top: 1.5rem; padding-bottom: 3rem; }}
    .hero {{ background: {CIANO}; color: white; padding: 28px 32px; border-radius: 18px; margin-bottom: 18px; }}
    .hero-title {{ font-size: 30px; font-weight: 800; margin-bottom: 5px; }}
    .hero-subtitle {{ font-size: 15px; opacity: .85; }}
    .metric-card {{ border-radius: 16px; padding: 20px 22px; min-height: 125px; box-shadow: 0 2px 8px rgba(0,0,0,.03); }}
    .metric-label {{ color: #687385; font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; }}
    .metric-value {{ font-size: 28px; font-weight: 800; margin-top: 7px; }}
    .metric-note {{ color: #7C8798; font-size: 12px; margin-top: 4px; }}
    .section-title {{ color: {TITULO_COR}; font-size: 18px; font-weight: 800; margin: 20px 0 8px 0; }}
    .chart-label {{ font-size: 15px; font-weight: 700; margin-bottom: 6px; }}
    .card-title {{ font-size: 20px; font-weight: 800; margin: 2px 0 12px 0; color: #1F2430; }}
    .big-number-sm {{ font-size: 26px; font-weight: 800; line-height: 1.2; }}
    .small-muted {{ color: #687385; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em; }}
    .regra-negocio {{ background: #FCEAF1; border-left: 4px solid {PINK}; border-radius: 8px; padding: 10px 14px; margin-top: 10px; font-size: 13px; color: #4A4A4A; }}
    div[data-testid="stDataFrame"] {{ border-radius: 12px; overflow: hidden; }}
    .stButton > button {{ border-radius: 10px; background-color: {PRIMARIO}; color: white; border: none; }}
</style>
""", unsafe_allow_html=True)

# ---------- Conexão com Google Sheets ----------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Nomes das abas na planilha do Google Sheets (devem ser idênticos aos da planilha Excel original)
ABAS = ["gastos", "fixos", "cartao", "entradas", "investimentos", "contas_bancarias", "config"]

# Colunas que devem ser tratadas como número em cada aba
COLUNAS_NUMERICAS = {
    "gastos": ["valor"],
    "fixos": ["valor"],
    "cartao": ["parcelas", "parcela_atual", "valor_parcela"],
    "entradas": ["valor"],
    "investimentos": ["percentual_salario", "valor_meta", "valor_investido"],
    "contas_bancarias": ["saldo_inicial", "saldo_final"],
    "config": ["limite_categoria"],
}

# Colunas booleanas (na planilha, digite TRUE/FALSE ou VERDADEIRO/FALSO nessas células,
# ou use uma caixa de seleção — checkbox — do próprio Google Sheets)
COLUNAS_BOOLEANAS = {
    "fixos": ["pago"],
    "cartao": ["pago"],
}

# Colunas de data (devem estar formatadas como Data no Google Sheets)
COLUNAS_DATA = {
    "gastos": ["data_compra", "data_pagamento"],
    "fixos": ["vencimento"],
    "cartao": ["vencimento"],
    "entradas": ["data"],
}


def _para_numero(serie):
    # Com value_render_option="UNFORMATTED_VALUE", o Sheets já entrega o número
    # bruto (float), sem formatação regional — só garantimos o tipo numérico.
    return pd.to_numeric(serie, errors="coerce")


def _serial_para_data(serie):
    # Datas do Google Sheets (modo bruto) vêm como número de dias desde 30/12/1899,
    # o mesmo padrão usado pelo Excel. Convertemos esse número para data de verdade.
    numeros = pd.to_numeric(serie, errors="coerce")
    return pd.to_datetime(numeros, unit="D", origin="1899-12-30", errors="coerce")


def _para_booleano(serie):
    mapa = {"TRUE": True, "VERDADEIRO": True, "SIM": True, "1": True,
            "FALSE": False, "FALSO": False, "NAO": False, "NÃO": False, "0": False, "": False}
    return serie.astype(str).str.strip().str.upper().map(mapa).fillna(False)


@st.cache_data(ttl=300, show_spinner="Atualizando dados da planilha...")
def carregar_planilhas():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    cliente = gspread.authorize(creds)
    planilha = cliente.open_by_key(st.secrets["sheet_id"])

    dados = {}
    for aba in ABAS:
        try:
            worksheet = planilha.worksheet(aba)
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"Não encontrei uma aba chamada '{aba}' na planilha do Google Sheets.")
            st.stop()
        registros = worksheet.get_all_records(value_render_option="UNFORMATTED_VALUE")
        df = pd.DataFrame(registros)

        for coluna in COLUNAS_NUMERICAS.get(aba, []):
            if coluna in df.columns:
                df[coluna] = _para_numero(df[coluna])

        for coluna in COLUNAS_DATA.get(aba, []):
            if coluna in df.columns:
                df[coluna] = _serial_para_data(df[coluna])

        for coluna in COLUNAS_BOOLEANAS.get(aba, []):
            if coluna in df.columns:
                df[coluna] = _para_booleano(df[coluna])

        dados[aba] = df
    return dados


try:
    planilhas = carregar_planilhas()
except Exception as e:
    st.error(
        "Não consegui me conectar à planilha do Google Sheets. "
        "Verifique se os 'secrets' (sheet_id e gcp_service_account) estão configurados "
        "e se a planilha foi compartilhada com o e-mail da conta de serviço.\n\n"
        f"Detalhe técnico: {e}"
    )
    st.stop()

# --- Tratamento Inicial dos Dados ---
gastos = planilhas["gastos"].rename(columns={
    "data_compra": "Data compra", "data_pagamento": "Data pagamento",
    "descricao": "Gasto", "valor": "Valor", "categoria": "Categoria", "forma_pagamento": "Pagamento"
}).copy()
gastos["Data compra"] = pd.to_datetime(gastos["Data compra"], dayfirst=True, errors="coerce")
gastos["Data pagamento"] = pd.to_datetime(gastos["Data pagamento"], dayfirst=True, errors="coerce")

fixos = planilhas["fixos"].rename(columns={
    "nome": "Gasto", "valor": "Valor", "vencimento": "Vencimento",
    "categoria": "Categoria", "forma_pagamento": "Pagamento", "pago": "Pago"
}).copy()
fixos["Vencimento"] = pd.to_datetime(fixos["Vencimento"], dayfirst=True, errors="coerce")

cartao = planilhas["cartao"].rename(columns={
    "compra": "Compra", "cartao": "Cartão", "parcelas": "Parcelamento",
    "parcela_atual": "Parcela atual", "valor_parcela": "Valor parcela",
    "categoria": "Categoria", "vencimento": "Vencimento", "pago": "Pago"
}).copy()
cartao["Vencimento"] = pd.to_datetime(cartao["Vencimento"], dayfirst=True, errors="coerce")

entradas = planilhas["entradas"].rename(columns={
    "data": "Data", "descricao": "Descrição", "valor": "Valor"
}).copy()
entradas["Data"] = pd.to_datetime(entradas["Data"], dayfirst=True, errors="coerce")

investimentos = planilhas["investimentos"].rename(columns={
    "mes": "Mês", "objetivo": "Objetivo", "percentual_salario": "% salário",
    "valor_meta": "Meta", "valor_investido": "Investido", "realizado": "Realizado"
}).copy()

contas = planilhas["contas_bancarias"].rename(columns={
    "conta": "Conta", "banco": "Banco",
    "saldo_inicial": "Saldo Inicial", "saldo_final": "Saldo Final"
}).copy()

config = planilhas["config"].copy()

LIMITES_CATEGORIA = dict(zip(
    config.loc[config["limite_categoria"].notna(), "categoria"],
    config.loc[config["limite_categoria"].notna(), "limite_categoria"]
)) if "limite_categoria" in config.columns else {}
_param = config.loc[config.get("parametro") == "limite_mensal_total", "valor"] if "parametro" in config.columns else pd.Series(dtype=float)
LIMITE_MENSAL = float(_param.iloc[0]) if not _param.empty else np.inf


def atribuir_cores_categoria(categorias, status_lista, paleta, cor_alerta):
    cores = []
    ultima_normal = None
    ponteiro = 0
    for status in status_lista:
        if status == "Acima do limite":
            cores.append(cor_alerta)
            continue
        candidata = paleta[ponteiro % len(paleta)]
        if candidata == ultima_normal:
            ponteiro += 1
            candidata = paleta[ponteiro % len(paleta)]
        cores.append(candidata)
        ultima_normal = candidata
        ponteiro += 1
    return cores


def fmt_moeda(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def grafico_barra(df, campo_categoria, campo_valor, cores, dominio=None, tooltip=None, altura=280):
    valor_max = df[campo_valor].max() if not df.empty else 0
    escala_cor = alt.Scale(domain=dominio, range=cores) if dominio else alt.Scale(range=cores)
    return alt.Chart(df).mark_bar().encode(
        x=alt.X(
            f"{campo_valor}:Q", title="Valor gasto (R$)",
            scale=alt.Scale(domain=[0, valor_max * 1.15])
        ),
        y=alt.Y(
            f"{campo_categoria}:N", sort=None, title="",
            axis=alt.Axis(labelFontWeight="bold", labelFontSize=13)
        ),
        color=alt.Color(f"{campo_categoria}:N", scale=escala_cor, legend=None),
        tooltip=tooltip or [campo_categoria, campo_valor]
    ).properties(
        height=altura,
        background=FUNDO_GRAFICO,
        padding={"left": 5, "right": 25, "top": 5, "bottom": 5}
    )

# ---------- Header ----------
st.markdown("""
<div class="hero">
    <div class="hero-title"> Meu Controle Financeiro</div>
    <div class="hero-subtitle"> Painel de gastos pessoais.</div>
</div>
""", unsafe_allow_html=True)

# Botão manual para forçar atualização (limpa o cache e recarrega os dados agora)
col_refresh, _ = st.columns([1, 5])
with col_refresh:
    if st.button("🔄 Atualizar dados agora"):
        st.cache_data.clear()
        st.rerun()

# ---------- Filtros Globais (Ano e Mês) ----------
c1, c2, c3 = st.columns([1, 1, 2])

anos_disponiveis = sorted(gastos["Data compra"].dt.year.dropna().unique(), reverse=True)
with c1:
    ano_sel = st.selectbox("Ano", anos_disponiveis if len(anos_disponiveis) else [pd.Timestamp.now().year])

meses_map = {
    1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Março", 4: "04 - Abril",
    5: "05 - Maio", 6: "06 - Junho", 7: "07 - Julho", 8: "08 - Agosto",
    9: "09 - Setembro", 10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro"
}

# Filtrar os meses que contêm dados para o ano selecionado
meses_no_ano = sorted(gastos[gastos["Data compra"].dt.year == ano_sel]["Data compra"].dt.month.dropna().unique())
opcoes_meses = [meses_map[int(m)] for m in meses_no_ano] if len(meses_no_ano) else list(meses_map.values())

with c2:
    mes_sel_str = st.selectbox("Mês", opcoes_meses)
    mes_sel = int(mes_sel_str.split(" - ")[0])

with c3:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    st.caption("Os filtros de mês e ano controlam toda a página. Os dados são atualizados automaticamente a cada 5 minutos, ou na hora pelo botão acima.")

st.divider()

# ---------- Aplicação dos Filtros nos DataFrames ----------
gastos_m = gastos[(gastos["Data compra"].dt.year == ano_sel) & (gastos["Data compra"].dt.month == mes_sel)].copy()
entradas_m = entradas[(entradas["Data"].dt.year == ano_sel) & (entradas["Data"].dt.month == mes_sel)].copy()
fixos_m = fixos[(fixos["Vencimento"].dt.year == ano_sel) & (fixos["Vencimento"].dt.month == mes_sel)].copy()
cartao_m = cartao[(cartao["Vencimento"].dt.year == ano_sel) & (cartao["Vencimento"].dt.month == mes_sel)].copy()

# Filtro exato para a aba Investimentos (Formato 'YYYY-MM', ex: '2026-08')
chave_mes_sel = f"{ano_sel}-{mes_sel:02d}"

if "Mês" in investimentos.columns:
    investimentos["mes_str"] = pd.to_datetime(investimentos["Mês"], errors="coerce").dt.strftime("%Y-%m")
    investimentos["mes_str"] = investimentos["mes_str"].fillna(investimentos["Mês"].astype(str).str.strip())
    investimentos_m = investimentos[investimentos["mes_str"] == chave_mes_sel].copy()
else:
    investimentos_m = investimentos.copy()

# Filtro para Contas Bancárias (se houver ano e mes)
if "ano" in contas.columns and "mes" in contas.columns:
    contas_m = contas[(contas["ano"] == ano_sel) & (contas["mes"] == mes_sel)].copy()
elif "mes" in contas.columns:
    contas["mes_str"] = pd.to_datetime(contas["mes"], errors="coerce").dt.strftime("%Y-%m")
    contas["mes_str"] = contas["mes_str"].fillna(contas["mes"].astype(str).str.strip())
    contas_m = contas[contas["mes_str"] == chave_mes_sel].copy()
else:
    contas_m = contas.copy()

# ---------- Main metrics ----------
total_gastos_variaveis = gastos_m["Valor"].sum()
total_fixos = fixos_m["Valor"].sum()
total_cartao = cartao_m["Valor parcela"].sum()

total_gastos = total_gastos_variaveis + total_fixos + total_cartao

total_entradas = entradas_m["Valor"].sum()
total_investido = investimentos_m["Investido"].sum() if not investimentos_m.empty else 0.0

saldo = total_entradas - total_gastos - total_investido 

gastos_excedeu = total_gastos > LIMITE_MENSAL

metrics = [
    {
        "label": "GASTOS DO MÊS",
        "valor": fmt_moeda(total_gastos),
        "nota": "Acima do limite mensal de " + fmt_moeda(LIMITE_MENSAL) if gastos_excedeu else "Dentro do limite mensal definido",
        "cor": ALERTA if gastos_excedeu else PALETA_METRICAS[0],
    },
    {"label": "ENTRADAS", "valor": fmt_moeda(total_entradas), "nota": "Salário", "cor": PALETA_METRICAS[1]},
    {"label": "INVESTIDO", "valor": fmt_moeda(total_investido), "nota": "Conforme metas do mês", "cor": PALETA_METRICAS[2]},
    {"label": "DISPONÍVEL", "valor": fmt_moeda(saldo), "nota": "Entradas menos gastos do mês", "cor": PALETA_METRICAS[3]},
]
cols = st.columns(4)
for col, m in zip(cols, metrics):
    with col:
        fundo = "#FCEAF1" if m["cor"] == ALERTA else "white"
        st.markdown(f"""
        <div class="metric-card" style="background:{fundo}; border:1px solid #E2E7EF; border-top:3px solid {m['cor']};">
            <div class="metric-label">{m['label']}</div>
            <div class="metric-value" style="color:{m['cor']};">{m['valor']}</div>
            <div class="metric-note">{m['nota']}</div>
        </div>
        """, unsafe_allow_html=True)

# ---------- Charts: para onde foi o dinheiro ----------
st.markdown('<div class="section-title">📊 Para onde foi meu dinheiro?</div>', unsafe_allow_html=True)
left, right = st.columns(2)

with left:
    with st.container(border=True):
        st.markdown(f'<div class="chart-label" style="color:{VINHO};">Visão geral por categoria</div>', unsafe_allow_html=True)

        if not gastos_m.empty:
            cat_df = gastos_m.groupby("Categoria", as_index=False)["Valor"].sum()
            cat_df["Limite"] = cat_df["Categoria"].map(LIMITES_CATEGORIA).fillna(np.inf)
            cat_df["Status"] = np.where(cat_df["Valor"] > cat_df["Limite"], "Acima do limite", "Dentro do limite")
            cat_df = cat_df.sort_values("Valor", ascending=True).reset_index(drop=True)
            cat_df["Cor"] = atribuir_cores_categoria(cat_df["Categoria"].tolist(), cat_df["Status"].tolist(), PALETA_CATEGORIAS, ALERTA)

            chart_categoria = grafico_barra(
                cat_df, "Categoria", "Valor",
                cores=cat_df["Cor"].tolist(),
                dominio=cat_df["Categoria"].tolist(),
                tooltip=["Categoria", "Valor", "Limite", "Status"]
            )
            st.altair_chart(chart_categoria, use_container_width=True)

            categorias_estouradas = cat_df.loc[cat_df["Status"] == "Acima do limite", "Categoria"].tolist()
            if categorias_estouradas:
                aviso = f"Neste mês, a(s) categoria(s) <b>{', '.join(categorias_estouradas)}</b> ultrapassaram o limite definido."
            else:
                aviso = "Neste mês, nenhuma categoria ultrapassou o limite definido."
        else:
            st.info("Sem registros de gastos para o mês selecionado.")
            aviso = "Sem dados no mês."

        st.markdown(f"""
        <div class="regra-negocio">
            <b>Aviso: {aviso}
        </div>
        """, unsafe_allow_html=True)

with right:
    with st.container(border=True):
        st.markdown(f'<div class="chart-label" style="color:{CIANO};">Visão geral por forma de pagamento</div>', unsafe_allow_html=True)

        if not gastos_m.empty:
            pag_df = gastos_m.groupby("Pagamento", as_index=False)["Valor"].sum().sort_values("Valor", ascending=True)

            chart_pagamento = grafico_barra(
                pag_df, "Pagamento", "Valor",
                cores=PALETA_PAGAMENTO,
                tooltip=["Pagamento", "Valor"]
            )
            st.altair_chart(chart_pagamento, use_container_width=True)
        else:
            st.info("Sem registros para o mês selecionado.")

# ---------- Credit + fixed ----------
st.markdown('<div class="section-title">💳 Cartão e gastos fixos</div>', unsafe_allow_html=True)

with st.container(border=True):
    st.markdown('<div class="card-title">Cartão de crédito</div>', unsafe_allow_html=True)
    st.markdown('<div class="small-muted">Total do mês</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="big-number-sm" style="color:{VINHO};">{fmt_moeda(total_cartao)}</div>', unsafe_allow_html=True)
    st.caption("Total das parcelas consideradas no mês.")

    display_cartao = cartao_m[["Compra", "Cartão", "Parcelamento", "Parcela atual", "Valor parcela", "Categoria", "Vencimento", "Pago"]].copy()
    if not display_cartao.empty:
        display_cartao["Vencimento"] = display_cartao["Vencimento"].dt.strftime("%d/%m/%Y")
        display_cartao["Pago"] = display_cartao["Pago"].map({True: "✅", False: "⬜"})
    st.dataframe(
        display_cartao,
        use_container_width=True,
        hide_index=True,
        column_config={"Valor parcela": st.column_config.NumberColumn(format="R$ %.2f")}
    )

with st.container(border=True):
    st.markdown('<div class="card-title">Gastos fixos</div>', unsafe_allow_html=True)
    col_pago, col_pendente = st.columns(2)
    fixos_pagos = fixos_m.loc[fixos_m["Pago"] == True, "Valor"].sum()
    fixos_pendentes = fixos_m.loc[fixos_m["Pago"] == False, "Valor"].sum()

    with col_pago:
        st.markdown('<div class="small-muted">Pago</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="big-number-sm" style="color:{CIANO};">{fmt_moeda(fixos_pagos)}</div>', unsafe_allow_html=True)
    with col_pendente:
        st.markdown('<div class="small-muted">Pendente</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="big-number-sm" style="color:{PINK};">{fmt_moeda(fixos_pendentes)}</div>', unsafe_allow_html=True)
    st.caption("Resumo dos pagamentos fixos do mês.")

    display_fixos = fixos_m[["Gasto", "Valor", "Categoria", "Pagamento", "Vencimento", "Pago"]].copy()
    if not display_fixos.empty:
        display_fixos["Vencimento"] = display_fixos["Vencimento"].dt.strftime("%d/%m/%Y")
        display_fixos["Pago"] = display_fixos["Pago"].map({True: "✅", False: "⬜"})
    st.dataframe(
        display_fixos,
        use_container_width=True,
        hide_index=True,
        column_config={"Valor": st.column_config.NumberColumn(format="R$ %.2f")}
    )

# ---------- Investments ----------
st.markdown('<div class="section-title">📈 Investimentos</div>', unsafe_allow_html=True)
if not investimentos_m.empty:
    inv_cols = st.columns(len(investimentos_m))
    for idx, (col, (_, row)) in enumerate(zip(inv_cols, investimentos_m.iterrows())):
        percentual = min(row["Investido"] / row["Meta"], 1) if row["Meta"] else 0
        cor_meta = PALETA_CONTAS[idx % len(PALETA_CONTAS)]
        with col:
            with st.container(border=True):
                icone = "✈️" if "viagem" in str(row["Objetivo"]).lower() else "🏠"
                st.markdown(f'<div class="card-title" style="color:{cor_meta};">{icone} {row["Objetivo"]}</div>', unsafe_allow_html=True)
                st.write(f"**Meta:** {fmt_moeda(row['Meta'])}")
                st.write(f"**Investido:** {fmt_moeda(row['Investido'])}")
                st.progress(percentual)
                st.caption(f"{percentual:.0%} da meta mensal")
else:
    st.info("Nenhum investimento registrado para o mês selecionado.")

# ---------- Detail ----------
st.markdown('<div class="section-title">📋 Detalhes dos gastos</div>', unsafe_allow_html=True)

f1, f2 = st.columns(2)
with f1:
    categorias = st.multiselect("Filtrar categoria", sorted(gastos_m["Categoria"].unique()) if not gastos_m.empty else [])
with f2:
    pagamentos = st.multiselect("Filtrar pagamento", sorted(gastos_m["Pagamento"].unique()) if not gastos_m.empty else [])

detalhes = gastos_m.copy()
if categorias:
    detalhes = detalhes[detalhes["Categoria"].isin(categorias)]
if pagamentos:
    detalhes = detalhes[detalhes["Pagamento"].isin(pagamentos)]

detalhes_exibicao = detalhes[["Data compra", "Data pagamento", "Gasto", "Valor", "Categoria", "Pagamento"]].copy()
if not detalhes_exibicao.empty:
    detalhes_exibicao["Data compra"] = detalhes_exibicao["Data compra"].dt.strftime("%d/%m/%Y")
    detalhes_exibicao["Data pagamento"] = detalhes_exibicao["Data pagamento"].dt.strftime("%d/%m/%Y")

st.dataframe(
    detalhes_exibicao,
    use_container_width=True,
    hide_index=True,
    column_config={"Valor": st.column_config.NumberColumn(format="R$ %.2f")}
)

# ---------- Evolution ----------
st.markdown('<div class="section-title">📈 Evolução mensal</div>', unsafe_allow_html=True)
evolucao = gastos.groupby(gastos["Data compra"].dt.to_period("M"))["Valor"].sum().reset_index()
evolucao["Mês"] = evolucao["Data compra"].dt.strftime("%b/%Y")
evolucao = evolucao[["Mês", "Valor"]].rename(columns={"Valor": "Gastos"})

chart_evolucao = alt.Chart(evolucao).mark_line(point=True, color=AZUL_LILAS, strokeWidth=3).encode(
    x=alt.X("Mês:N", title="", sort=None),
    y=alt.Y("Gastos:Q", title="Valor gasto (R$)"),
    tooltip=["Mês", "Gastos"]
).properties(
    height=260,
    background=FUNDO_GRAFICO,
    padding={"left": 5, "right": 25, "top": 5, "bottom": 5}
)

with st.container(border=True):
    st.altair_chart(chart_evolucao, use_container_width=True)
    if len(evolucao) < 2:
        st.caption("À medida que novos meses forem lançados na aba 'gastos' da planilha, a evolução aparecerá aqui.")

# ---------- Saldo em contas bancárias ----------
st.markdown('<div class="section-title">🏦 Saldo em contas bancárias</div>', unsafe_allow_html=True)

total_saldo_inicial = contas_m["Saldo Inicial"].sum() if "Saldo Inicial" in contas_m.columns and not contas_m.empty else 0.0
total_saldo_final = contas_m["Saldo Final"].sum() if "Saldo Final" in contas_m.columns and not contas_m.empty else 0.0

with st.container(border=True):
    col_ini_tot, col_fim_tot = st.columns(2)
    with col_ini_tot:
        st.markdown('<div class="small-muted">Saldo Inicial Total</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="big-number-sm" style="color:{CIANO};">{fmt_moeda(total_saldo_inicial)}</div>', unsafe_allow_html=True)
    with col_fim_tot:
        st.markdown('<div class="small-muted">Saldo Final Total</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="big-number-sm" style="color:{VINHO};">{fmt_moeda(total_saldo_final)}</div>', unsafe_allow_html=True)

if not contas_m.empty:
    conta_cols = st.columns(len(contas_m))
    for idx, (col, (_, row)) in enumerate(zip(conta_cols, contas_m.iterrows())):
        cor = PALETA_CONTAS[idx % len(PALETA_CONTAS)]
        with col:
            with st.container(border=True):
                st.markdown(f'<div class="chart-label" style="color:{cor};">🏦 {row["Conta"]}</div>', unsafe_allow_html=True)
                st.caption(row["Banco"])
                s_ini = fmt_moeda(row["Saldo Inicial"]) if "Saldo Inicial" in row else "R$ 0,00"
                s_fim = fmt_moeda(row["Saldo Final"]) if "Saldo Final" in row else "R$ 0,00"
                st.markdown(f"**Início:** {s_ini}")
                st.markdown(f"**Fim:** {s_fim}")
else:
    st.info("Nenhum registro de conta bancária para o mês selecionado.")

st.caption("Criado por eu rainha Keyla - atualize a planilha no Google Sheets para ver as mudanças (até 5 min de atraso, ou use o botão 'Atualizar dados agora').")
