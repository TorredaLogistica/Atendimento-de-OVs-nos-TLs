import os, re, unicodedata
from urllib.parse import quote
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Controle de Recebimento de Materiais", page_icon="📦", layout="wide")
BASE_XLSX = "Controle de Recebimento de Materiais.xlsx"

def norm(v):
    s = "" if pd.isna(v) else str(v)
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().upper().strip()

@st.cache_data(show_spinner="Carregando base de recebimentos...")
def load_data():
    if os.path.exists(BASE_XLSX):
        d = pd.read_excel(
            BASE_XLSX,
            sheet_name="Recebimento Manual",
            engine="openpyxl",
            usecols=[
                "ANO/MÊS_RECEBIMENTO", "Tipo_recebimento", "STATUS RECEBIMENTO",
                "SLA_RECEBIMENTO", "SLA_ENVIO", "CD_CORRIGIDO", "DATA_BASE",
                "CD", "EMPRESA", "PROCESSO", "ORIGEM", "NOTAFISCAL",
                "DATA RECEBIMENTO", "DATA LANÇAMENTO SAP"
            ]
        )
    else:
        st.error(f"Base não encontrada. Inclua o arquivo {BASE_XLSX} na mesma pasta do app.")
        st.stop()
    d.columns = [str(c).strip() for c in d.columns]
    for c in ["DATA_BASE", "DATA RECEBIMENTO", "DATA LANÇAMENTO SAP"]:
        d[c] = pd.to_datetime(d[c], errors="coerce")
    d["SLA_RECEBIMENTO"] = pd.to_numeric(d["SLA_RECEBIMENTO"], errors="coerce")
    # Mês de referência: coluna G da base (DATA_BASE).
    d["MES_REF"] = d["DATA_BASE"].dt.to_period("M").dt.to_timestamp()

    # Mantém todos os tipos de recebimento da base, inclusive Obsoletos.
    # O filtro lateral permite selecionar Transferência, Reversa Massivo ou Obsoletos.
    d["FAIXA_SLA"] = pd.cut(d["SLA_RECEBIMENTO"], bins=[-float("inf"),0,1,2,3,4,float("inf")], labels=["D+0","D+1","D+2","D+3","D+4","Acima de D+4"])
    return d

def fmt_int(v): return f"{int(v):,}".replace(",", ".")
def fmt_pct(v): return f"{v:.1%}".replace(".", ",")

def metric_cards(data):
    total=len(data)
    buckets=[("D+0",0),("D+1",1),("D+2",2),("D+3",3),("D+4",4),("Acima de D+4",None)]
    cols=st.columns(6)
    for col,(label,lim) in zip(cols,buckets):
        if lim is None:
            qtd=int((data["SLA_RECEBIMENTO"]>4).sum()); pct=qtd/total if total else 0
            sub=f"{fmt_pct(pct)} do total"
        else:
            qtd=int((data["SLA_RECEBIMENTO"]==lim).sum())
            acum=int((data["SLA_RECEBIMENTO"]<=lim).sum())
            pct=acum/total if total else 0
            sub=f"Acumulado até D+{lim}: {fmt_pct(pct)}"
        col.metric(f"SLA {label}", fmt_int(qtd), sub)

def executive_summary(data, months=3):
    if data.empty: return "Sem dados para os filtros selecionados."
    mx=data["MES_REF"].max(); start=mx-pd.DateOffset(months=months-1)
    x=data[data["MES_REF"].between(start,mx)].copy()
    lines=["RESUMO EXECUTIVO | CONTROLE DE RECEBIMENTO DE MATERIAIS", ""]
    for mes,g in x.groupby("MES_REF", sort=True):
        total=len(g); ate1=(g["SLA_RECEBIMENTO"]<=1).sum(); ate4=(g["SLA_RECEBIMENTO"]<=4).sum(); fora=(g["SLA_RECEBIMENTO"]>4).sum()
        lines.append(f"{mes.strftime('%m/%Y')}: {fmt_int(total)} recebimentos | Até D+1: {fmt_pct(ate1/total if total else 0)} | Até D+4: {fmt_pct(ate4/total if total else 0)} | Acima de D+4: {fmt_int(fora)} ({fmt_pct(fora/total if total else 0)})")
    lines += ["", "Escopo: todos os tipos de recebimento disponíveis na base, conforme os filtros selecionados."]
    return "\n".join(lines)

st.markdown("""<style>
.block-container{padding-top:1.4rem}.kpi-note{background:#fff7f7;border-left:7px solid #e30613;padding:17px 22px;border-radius:10px;margin:8px 0 20px}.small{color:#65707d;font-size:.89rem}
</style>""", unsafe_allow_html=True)
st.title("📦 Controle de Recebimento de Materiais")
st.markdown('<div class="kpi-note"><b>Controle de Recebimentos</b><br>Monitoramento do prazo entre o recebimento físico e o lançamento no SAP, considerando os tipos selecionados nos filtros.</div>', unsafe_allow_html=True)

df=load_data()
with st.sidebar:
    st.header("Visualização")
    vis=st.radio("Selecione", ["📅 Visão diária", "📊 Evolução mensal"], label_visibility="collapsed")
    st.header("Filtros")
    meses=sorted(df["MES_REF"].dropna().unique(), reverse=True)
    mes=st.selectbox("Mês de referência", meses, format_func=lambda x: pd.Timestamp(x).strftime("%m/%Y"))
    def multi(label,col):
        vals=sorted(df[col].dropna().astype(str).str.strip().unique())
        return st.multiselect(label, vals)
    f_tipo=multi("Tipo de recebimento","Tipo_recebimento")
    f_cd=multi("CD corrigido","CD_CORRIGIDO")
    f_empresa=multi("Empresa","EMPRESA")
    f_processo=multi("Processo","PROCESSO")
    f_origem=multi("Origem","ORIGEM")
    nf=st.text_input("Buscar Nota Fiscal")

def apply(d):
    x=d.copy()
    for vals,col in [(f_tipo,"Tipo_recebimento"),(f_cd,"CD_CORRIGIDO"),(f_empresa,"EMPRESA"),(f_processo,"PROCESSO"),(f_origem,"ORIGEM")]:
        if vals: x=x[x[col].astype(str).str.strip().isin(vals)]
    if nf: x=x[x["NOTAFISCAL"].astype(str).str.contains(re.escape(nf),case=False,na=False)]
    return x
base=apply(df)

if vis.startswith("📅"):
    data=base[base["MES_REF"]==pd.Timestamp(mes)].copy()
    metric_cards(data)
    c1,c2=st.columns([1.55,1])
    with c1:
        daily=data.groupby(data["DATA RECEBIMENTO"].dt.date,as_index=False).agg(Recebimentos=("NOTAFISCAL","size"),SLA_ate_D1=("SLA_RECEBIMENTO",lambda s:(s<=1).mean()),SLA_ate_D4=("SLA_RECEBIMENTO",lambda s:(s<=4).mean()))
        fig=px.bar(daily,x="DATA RECEBIMENTO",y="Recebimentos",text_auto=True,title=f"Recebimentos por dia | {pd.Timestamp(mes).strftime('%m/%Y')}")
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        dist=data["FAIXA_SLA"].value_counts(sort=False).rename_axis("Faixa").reset_index(name="Quantidade")
        fig=px.bar(dist,x="Faixa",y="Quantidade",color="Faixa",text_auto=True,title="Distribuição por faixa de SLA")
        st.plotly_chart(fig,use_container_width=True)
    st.subheader("Detalhamento dos recebimentos")
    detail_cols=["DATA RECEBIMENTO","DATA LANÇAMENTO SAP","Tipo_recebimento","SLA_RECEBIMENTO","STATUS RECEBIMENTO","CD_CORRIGIDO","EMPRESA","PROCESSO","ORIGEM","NOTAFISCAL"]
    st.dataframe(data[detail_cols].sort_values("DATA RECEBIMENTO",ascending=False),use_container_width=True,hide_index=True)
else:
    qtd_meses=st.segmented_control("Período para comparação",[3,6,9,12],default=3,format_func=lambda n:f"Últimos {n} meses")
    maxm=base["MES_REF"].max()
    evo=base[base["MES_REF"]>=maxm-pd.DateOffset(months=int(qtd_meses)-1)].copy() if pd.notna(maxm) else base
    mensal=evo.groupby("MES_REF",as_index=False).agg(Recebimentos=("NOTAFISCAL","size"),SLA_D0=("SLA_RECEBIMENTO",lambda s:(s<=0).mean()),SLA_D1=("SLA_RECEBIMENTO",lambda s:(s<=1).mean()),SLA_D2=("SLA_RECEBIMENTO",lambda s:(s<=2).mean()),SLA_D3=("SLA_RECEBIMENTO",lambda s:(s<=3).mean()),SLA_D4=("SLA_RECEBIMENTO",lambda s:(s<=4).mean()))
    longa=mensal.melt(id_vars=["MES_REF","Recebimentos"],value_vars=["SLA_D0","SLA_D1","SLA_D2","SLA_D3","SLA_D4"],var_name="SLA",value_name="Percentual")
    longa["Mês"]=longa["MES_REF"].dt.strftime("%m/%Y")
    fig=px.line(longa,x="Mês",y="Percentual",color="SLA",markers=True,text=longa["Percentual"].map(lambda x:fmt_pct(x)),title="Evolução mensal do SLA acumulado")
    fig.update_yaxes(tickformat=".0%",range=[0,1.08]); fig.update_traces(textposition="top center")
    st.plotly_chart(fig,use_container_width=True)
    fig2=px.bar(mensal.assign(Mês=mensal["MES_REF"].dt.strftime("%m/%Y")),x="Mês",y="Recebimentos",text_auto=True,title="Volume mensal de recebimentos")
    st.plotly_chart(fig2,use_container_width=True)
    # Na tabela, converte a proporção decimal para percentual real.
    # Exemplo: 0.2708 passa a ser exibido como 27,08%, e não 0,3%.
    colunas_sla = ["SLA_D0", "SLA_D1", "SLA_D2", "SLA_D3", "SLA_D4"]
    tabela_mensal = mensal.assign(
        **{"Mês": mensal["MES_REF"].dt.strftime("%m/%Y")}
    )[["Mês", "Recebimentos"] + colunas_sla].copy()

    tabela_mensal[colunas_sla] = tabela_mensal[colunas_sla].mul(100)

    st.dataframe(
        tabela_mensal,
        use_container_width=True,
        hide_index=True,
        column_config={
            c: st.column_config.ProgressColumn(
                c,
                format="%.2f%%",
                min_value=0,
                max_value=100,
            )
            for c in colunas_sla
        },
    )

st.divider(); st.subheader("✉️ Envio por e-mail | resumo executivo dos últimos 3 meses")
resumo=executive_summary(base,3)
with st.expander("Visualizar resumo",expanded=False): st.text(resumo)
c1,c2=st.columns([2,1])
with c1:
    destinatarios=st.text_input("Destinatários",placeholder="nome@empresa.com; outro@empresa.com")
with c2:
    assunto=st.text_input("Assunto",value="Controle de Recebimento de Materiais | Resumo dos últimos 3 meses")
url="https://outlook.office.com/mail/deeplink/compose?to="+quote(destinatarios.replace(";",","))+"&subject="+quote(assunto)+"&body="+quote(resumo)
st.link_button("Abrir e-mail no Outlook Web",url,use_container_width=True)
st.download_button("Baixar resumo em texto",resumo,file_name="Resumo_Executivo_Recebimentos.txt",mime="text/plain")
st.caption("O envio permanece manual e seguro: o botão abre a mensagem pronta no Outlook Web para revisão antes do envio.")
