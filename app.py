from __future__ import annotations

from io import BytesIO
from pathlib import Path
from datetime import date
import re
import unicodedata

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Indicador de Atendimento de OVs nos TLs", page_icon="📦", layout="wide")

ABA_PADRAO = "BASE OVS"
NOME_ARQUIVO_PADRAO = "Base OVs TLs.xlsx"


def normalizar_texto(valor: object) -> str:
    texto = "" if pd.isna(valor) else str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto).upper()


def localizar_coluna(colunas, candidatos, obrigatoria=True):
    mapa = {normalizar_texto(c): c for c in colunas}
    for candidato in candidatos:
        chave = normalizar_texto(candidato)
        if chave in mapa:
            return mapa[chave]
    if obrigatoria:
        raise ValueError(f"Coluna obrigatória não encontrada: {candidatos[0]}")
    return None


@st.cache_data(show_spinner=False)
def carregar_excel(conteudo: bytes, aba: str) -> pd.DataFrame:
    return pd.read_excel(BytesIO(conteudo), sheet_name=aba, engine="openpyxl", dtype=object)


def preparar_base(df: pd.DataFrame, data_referencia: pd.Timestamp) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    c_doc = localizar_coluna(df.columns, ["Doc. SD", "Documento SD", "OV"])
    c_pedido = localizar_coluna(df.columns, ["Data ped.", "Data pedido", "Data Pedido"])
    c_atendido = localizar_coluna(df.columns, ["Data atend.", "Data atendimento", "Data Atendida"])
    c_situacao = localizar_coluna(df.columns, ["Situação", "Situacao", "Status"])
    c_dif = localizar_coluna(df.columns, ["Dif. Dias", "Dif Dias", "Dias"], obrigatoria=False)

    for coluna in [c_pedido, c_atendido]:
        df[coluna] = pd.to_datetime(df[coluna], errors="coerce", dayfirst=True)

    # Garante uma OV por linha no indicador, evitando duplicidade acidental da base.
    df = df.dropna(subset=[c_doc]).drop_duplicates(subset=[c_doc], keep="last")

    df["Situação Normalizada"] = df[c_situacao].map(normalizar_texto)
    df["Pedido Atendido"] = df[c_atendido].notna()
    df["Pedido em Atendimento"] = ~df["Pedido Atendido"]

    if c_dif:
        dias_informados = pd.to_numeric(df[c_dif], errors="coerce")
    else:
        dias_informados = pd.Series(pd.NA, index=df.index, dtype="Float64")

    dias_calculados = (df[c_atendido].dt.normalize() - df[c_pedido].dt.normalize()).dt.days
    df["Dias para Atendimento"] = dias_informados.where(dias_informados.notna(), dias_calculados)

    # Contagem em dias úteis (segunda a sexta-feira), sem considerar feriados.
    data_final = data_referencia.normalize().to_datetime64().astype("datetime64[D]")
    dias_uteis = pd.Series(pd.NA, index=df.index, dtype="Int64")
    datas_validas = df[c_pedido].notna()
    if datas_validas.any():
        datas_inicio = (
            df.loc[datas_validas, c_pedido]
            .dt.normalize()
            .values.astype("datetime64[D]")
        )
        dias_uteis.loc[datas_validas] = np.busday_count(datas_inicio, data_final)
    df["Dias Úteis em Aberto"] = dias_uteis.clip(lower=0)
    df["Dias em Aberto"] = df["Dias Úteis em Aberto"]

    df["Faixa do Indicador"] = "Em atendimento"
    df.loc[df["Pedido Atendido"] & (df["Dias para Atendimento"] == 1), "Faixa do Indicador"] = "Atendido em D+1"
    df.loc[df["Pedido Atendido"] & (df["Dias para Atendimento"] == 2), "Faixa do Indicador"] = "Atendido em D+2"
    df.loc[df["Pedido Atendido"] & (df["Dias para Atendimento"] > 2), "Faixa do Indicador"] = "Atendido acima de D+2"
    df.loc[df["Pedido Atendido"] & (df["Dias para Atendimento"] <= 0), "Faixa do Indicador"] = "Atendido em D+0"
    df["Em Atraso"] = df["Pedido em Atendimento"] & (df["Dias Úteis em Aberto"] > 1)
    df["No Prazo"] = df["Pedido em Atendimento"] & (df["Dias Úteis em Aberto"] <= 1)

    return df


def lista_opcoes(df, coluna):
    if coluna not in df.columns:
        return []
    return sorted(df[coluna].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())


def excel_para_download(df: pd.DataFrame) -> bytes:
    saida = BytesIO()
    exportar = df.copy()
    for coluna in exportar.columns:
        if pd.api.types.is_datetime64_any_dtype(exportar[coluna]):
            exportar[coluna] = exportar[coluna].dt.date
    with pd.ExcelWriter(saida, engine="openpyxl") as writer:
        exportar.to_excel(writer, index=False, sheet_name="Dados Filtrados")
        ws = writer.book["Dados Filtrados"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True, color="FFFFFF")
            cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="C00000")
        for coluna in ws.columns:
            letra = coluna[0].column_letter
            maior = min(max(len(str(c.value or "")) for c in coluna) + 2, 45)
            ws.column_dimensions[letra].width = maior
    return saida.getvalue()


st.title("📦 Indicador de Atendimento de OVs nos TLs")
st.caption("Monitoramento dos pedidos atendidos por prazo e dos pedidos ainda em atendimento do Cnal de TELEVENDAS")

CAMINHO_BASE = Path(__file__).resolve().parent / NOME_ARQUIVO_PADRAO

with st.sidebar:
    st.header("Base de dados")
    st.caption(f"Fonte automática: {NOME_ARQUIVO_PADRAO}")

if not CAMINHO_BASE.exists():
    st.error(
        f'O arquivo "{NOME_ARQUIVO_PADRAO}" não foi encontrado no repositório. '
        'Salve a planilha na mesma pasta do arquivo app.py.'
    )
    st.stop()

try:
    tamanho_base = CAMINHO_BASE.stat().st_size
    if tamanho_base < 10_000:
        inicio_arquivo = CAMINHO_BASE.read_bytes()[:200]
        if b"git-lfs.github.com/spec" in inicio_arquivo:
            st.error(
                "A base foi encontrada apenas como ponteiro do Git LFS. "
                "O ambiente não baixou o conteúdo real do arquivo Excel."
            )
            st.stop()
    conteudo = CAMINHO_BASE.read_bytes()
    nome_arquivo = CAMINHO_BASE.name
except Exception as erro:
    st.error(f"Não foi possível acessar a base armazenada no repositório: {erro}")
    st.stop()

try:
    df_original = carregar_excel(conteudo, ABA_PADRAO)
except ValueError:
    df_original = carregar_excel(conteudo, 0)
except Exception as erro:
    st.error(f"Não foi possível ler o arquivo: {erro}")
    st.stop()

c_pedido = localizar_coluna(df_original.columns, ["Data ped.", "Data pedido", "Data Pedido"])
datas_pedido = pd.to_datetime(df_original[c_pedido], errors="coerce", dayfirst=True)
if datas_pedido.notna().sum() == 0:
    st.error("A coluna de Data do pedido não possui nenhuma data válida.")
    st.stop()
data_min = datas_pedido.min().date()
data_max = datas_pedido.max().date()

MESES_REFERENCIA = {
    "Todos os meses": None,
    "Janeiro": 1,
    "Fevereiro": 2,
    "Março": 3,
    "Abril": 4,
    "Maio": 5,
    "Junho": 6,
    "Julho": 7,
    "Agosto": 8,
    "Setembro": 9,
    "Outubro": 10,
    "Novembro": 11,
    "Dezembro": 12,
}

with st.sidebar:
    st.header("Filtros")
    mes_referencia = st.selectbox(
        "Mês de referência",
        options=list(MESES_REFERENCIA.keys()),
        index=0,
        help="O mês é aplicado sobre a Data do pedido.",
    )
    busca = st.text_input(
        "Buscar OV",
        placeholder="Digite o número da OV",
        help="A busca é realizada somente na coluna Doc. SD.",
    )
    data_pedido_filtro = st.date_input(
        "Data do pedido",
        value=None,
        min_value=data_min,
        max_value=max(data_max, date.today()),
        format="DD/MM/YYYY",
        help="Selecione uma data para visualizar somente os pedidos daquele dia. Deixe em branco para exibir todas as datas.",
    )

try:
    # A data de referência deixou de ser um filtro visível e passa a ser a data atual.
    df = preparar_base(df_original, pd.Timestamp.today().normalize())
except Exception as erro:
    st.error(f"Erro no tratamento da base: {erro}")
    st.stop()

# Filtros em cascata simples e objetivos.
with st.sidebar:
    selecoes = {}
    for coluna, titulo in [
        ("Reg.", "Região/UF"),
        ("Org. vendas", "Organização de vendas"),
        ("Situação", "Situação"),
    ]:
        opcoes = lista_opcoes(df, coluna)
        if opcoes:
            selecoes[coluna] = st.multiselect(titulo, opcoes)
    st.caption(f"Fonte automática: {nome_arquivo}")

filtrado = df.copy()
numero_mes = MESES_REFERENCIA[mes_referencia]
if numero_mes is not None:
    filtrado = filtrado[filtrado[c_pedido].dt.month == numero_mes]
if data_pedido_filtro is not None:
    data_selecionada = pd.Timestamp(data_pedido_filtro)
    filtrado = filtrado[filtrado[c_pedido].dt.normalize() == data_selecionada.normalize()]
for coluna, valores in selecoes.items():
    if valores:
        filtrado = filtrado[filtrado[coluna].astype(str).isin(valores)]
if busca.strip():
    c_doc_busca = localizar_coluna(filtrado.columns, ["Doc. SD", "Documento SD", "OV"])
    filtrado = filtrado[
        filtrado[c_doc_busca].astype(str).str.contains(
            busca.strip(), case=False, na=False, regex=False
        )
    ]

m_d0 = int(((filtrado["Pedido Atendido"]) & (filtrado["Dias para Atendimento"] == 0)).sum())
m_d1 = int(((filtrado["Pedido Atendido"]) & (filtrado["Dias para Atendimento"] == 1)).sum())
m_d2 = int(((filtrado["Pedido Atendido"]) & (filtrado["Dias para Atendimento"] == 2)).sum())
m_acima = int(((filtrado["Pedido Atendido"]) & (filtrado["Dias para Atendimento"] > 2)).sum())
m_em_atendimento = int(filtrado["Pedido em Atendimento"].sum())
m_no_prazo = int(filtrado["No Prazo"].sum())
m_atraso = int(filtrado["Em Atraso"].sum())

# Cards clicáveis: funcionam como botões para filtrar a tabela detalhada.
st.markdown("""
<style>
div[data-testid="stButton"] > button {
    width: 100%;
    min-height: 142px;
    border-radius: 14px;
    border: 2px solid #d7dce2;
    background: linear-gradient(145deg, #ffffff, #f4f6f8);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.10);
    font-size: 23px;
    font-weight: 700;
    white-space: pre-line;
    transition: all 0.15s ease-in-out;
}
div[data-testid="stButton"] > button:hover {
    border-color: #c00000;
    color: #c00000;
    transform: translateY(-2px);
    box-shadow: 0 7px 16px rgba(192, 0, 0, 0.18);
}
div[data-testid="stButton"] > button[kind="primary"] {
    border-color: #c00000;
    background: linear-gradient(145deg, #c00000, #940000);
    color: #ffffff;
}
</style>
""", unsafe_allow_html=True)

if "card_selecionado" not in st.session_state:
    st.session_state.card_selecionado = None

def alternar_card(nome):
    st.session_state.card_selecionado = None if st.session_state.card_selecionado == nome else nome

total_pedidos = len(filtrado)

def percentual_sobre_total(valor):
    return (valor / total_pedidos * 100) if total_pedidos else 0.0

cards = [
    ("d0", "SLA D+0", m_d0),
    ("d1", "SLA D+1", m_d1),
    ("d2", "SLA D+2", m_d2),
    ("acima_d2", "SLA Acima de D+2", m_acima),
    ("em_atendimento", "Em Atendimento", m_em_atendimento),
    ("no_prazo", "No Prazo", m_no_prazo),
    ("em_atraso", "Em Atraso", m_atraso),
]
colunas_cards = st.columns(7)
for coluna_card, (chave, titulo, valor) in zip(colunas_cards, cards):
    selecionado = st.session_state.card_selecionado == chave
    percentual = percentual_sobre_total(valor)
    rotulo_card = (
        f"**{titulo}**\n\n"
        f"{valor:,}\n\n"
        f"{percentual:.1f}% do total"
    ).replace(",", "X").replace(".", ",").replace("X", ".")
    coluna_card.button(
        rotulo_card,
        key=f"card_{chave}",
        type="primary" if selecionado else "secondary",
        use_container_width=True,
        on_click=alternar_card,
        args=(chave,),
        help="Clique para filtrar a tabela. Clique novamente para remover o filtro.",
    )

card_ativo = st.session_state.card_selecionado
if card_ativo == "d0":
    tabela_filtrada = filtrado[filtrado["Pedido Atendido"] & (filtrado["Dias para Atendimento"] == 0)].copy()
elif card_ativo == "d1":
    tabela_filtrada = filtrado[filtrado["Pedido Atendido"] & (filtrado["Dias para Atendimento"] == 1)].copy()
elif card_ativo == "d2":
    tabela_filtrada = filtrado[filtrado["Pedido Atendido"] & (filtrado["Dias para Atendimento"] == 2)].copy()
elif card_ativo == "acima_d2":
    tabela_filtrada = filtrado[filtrado["Pedido Atendido"] & (filtrado["Dias para Atendimento"] > 2)].copy()
elif card_ativo == "em_atendimento":
    tabela_filtrada = filtrado[filtrado["Pedido em Atendimento"]].copy()
elif card_ativo == "no_prazo":
    tabela_filtrada = filtrado[filtrado["No Prazo"]].copy()
elif card_ativo == "em_atraso":
    tabela_filtrada = filtrado[filtrado["Em Atraso"]].copy()
else:
    tabela_filtrada = filtrado.copy()

st.divider()
titulo_ativo = next((titulo for chave, titulo, _ in cards if chave == card_ativo), "Todos os pedidos")
st.subheader(f"Detalhamento dos pedidos — {titulo_ativo}")
if card_ativo:
    st.caption("O card destacado está filtrando a tabela. Clique novamente no mesmo card para exibir todos os pedidos.")

ordem = ["No Prazo", "Em Atraso", "Faixa do Indicador", "Dias Úteis em Aberto", "Dias para Atendimento"]
colunas_exibicao = [c for c in ["Doc. SD", "Denominação", "Cliente", "Org. vendas", "Reg.", "Local", "Data ped.", "Data atend.", "Data exped.", "Nome", "Localidade", "CEP", "Data SIAKI", "Dif. Dias", "Situação"] if c in tabela_filtrada.columns] + ordem
resultado = tabela_filtrada[colunas_exibicao].sort_values(["Em Atraso", "Dias Úteis em Aberto"], ascending=[False, False])

st.dataframe(
    resultado,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Data ped.": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Data atend.": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Data exped.": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Data SIAKI": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "No Prazo": st.column_config.CheckboxColumn(),
        "Em Atraso": st.column_config.CheckboxColumn(),
    },
)

st.download_button(
    "⬇️ Extrair dados filtrados em Excel",
    data=excel_para_download(resultado),
    file_name=f"Indicador_OVs_{pd.Timestamp.today():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

with st.expander("Regras utilizadas no indicador"):
    st.markdown("""
- **SLA D+0:** pedido com data de atendimento preenchida no mesmo dia do pedido, com 0 dia em `Dif. Dias`.
- **Atendido em D+1:** pedido com data de atendimento preenchida e 1 dia em `Dif. Dias`.
- **Atendido em D+2:** pedido com data de atendimento preenchida e 2 dias em `Dif. Dias`.
- **Atendido acima de D+2:** pedido com data de atendimento preenchida e mais de 2 dias em `Dif. Dias`.
- **Em atendimento:** pedido sem data de atendimento preenchida.
- **No Prazo:** pedido ainda em atendimento que permanece em D+0 ou D+1 na data atual.
- **Em atraso:** pedido em atendimento com mais de D+1, calculado da data do pedido até a data atual.
- O cálculo considera **dias úteis (segunda a sexta-feira)**. Para pedidos atendidos, a classificação dos SLAs segue a coluna `Dif. Dias` da base.
""")
