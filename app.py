import io
import csv
import time
import statistics
import datetime as dt
from collections import defaultdict

import requests
import streamlit as st

st.set_page_config(
    page_title="OMIE Portugal - Precos Horarios",
    layout="centered",
)

st.title("OMIE Portugal - Precos Horarios + Analise BESS")
st.caption("Dados reais do mercado diario MIBEL - Fonte: www.omie.es")
st.divider()

# --- Datas ---
st.subheader("Periodo de dados")
col1, col2 = st.columns(2)
with col1:
    data_ini = st.date_input("Data de inicio", value=dt.date(2025, 1, 1),
                             min_value=dt.date(2015, 1, 1), max_value=dt.date.today(),
                             format="DD/MM/YYYY")
with col2:
    data_fim = st.date_input("Data de fim", value=dt.date(2025, 3, 31),
                             min_value=dt.date(2015, 1, 1), max_value=dt.date.today(),
                             format="DD/MM/YYYY")

if data_fim < data_ini:
    st.error("A data de fim tem de ser posterior a data de inicio.")
    st.stop()

n_dias = (data_fim - data_ini).days + 1
st.info(f"**{n_dias} dias** selecionados")
st.divider()

# --- Parametros BESS ---
st.subheader("Parametros BESS")
col3, col4, col5 = st.columns(3)
with col3:
    capacidade = st.number_input("Capacidade (MWh)", value=1.0, min_value=0.1, step=0.1)
    potencia   = st.number_input("Potencia (MW)",    value=0.5, min_value=0.1, step=0.1)
with col4:
    eficiencia = st.number_input("Eficiencia RT (%)", value=88, min_value=50, max_value=100) / 100
    opex       = st.number_input("OPEX (EUR/MWh/ano)", value=8, min_value=0, step=1)
with col5:
    capex     = st.number_input("CAPEX (EUR/kWh)",  value=250, min_value=0, step=10)
    vida_util = st.number_input("Vida util (anos)", value=15,  min_value=1, max_value=30)
    wacc      = st.number_input("WACC (%)",         value=7,   min_value=1, max_value=30) / 100

st.divider()

# --- Download OMIE ---
OMIE_URL = (
    "https://www.omie.es/pt/file-download"
    "?parents=marginalpdbcpt"
    "&filename=marginalpdbcpt_{date}.1"
)

def download_day(d):
    url = OMIE_URL.format(date=d.strftime("%Y%m%d"))
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if r.status_code != 200:
            return None
        records = []
        for line in r.text.splitlines():
            line = line.strip()
            if not line or not line[0].isdigit():
                continue
            parts = [p.strip().replace(",", ".") for p in line.split(";")]
            if len(parts) < 5:
                continue
            try:
                hour  = int(float(parts[3]))
                price = float(parts[4])
            except (ValueError, IndexError):
                continue
            h0   = hour - 1
            wday = d.weekday()
            records.append({
                "Data"     : str(d),
                "Hora"     : hour,
                "Horario"  : f"{h0:02d}:00-{hour:02d}:00",
                "Timestamp": f"{str(d)} {h0:02d}:00",
                "Preco_EUR_MWh": round(price, 2),
                "Mes"      : str(d)[:7],
                "DiaSemana": ["Seg","Ter","Qua","Qui","Sex","Sab","Dom"][wday % 7],
                "Periodo"  : (
                    "Vazio" if h0 in range(0, 7) else
                    "Cheia" if h0 in range(7, 10) or h0 in range(20, 24) else
                    "Ponta"
                ),
            })
        return records if records else None
    except Exception:
        return None

# --- Botao ---
if st.button("Descarregar dados e gerar CSV", type="primary", use_container_width=True):

    all_records = []
    falhas = []
    progress = st.progress(0, text="A iniciar download...")
    status   = st.empty()

    d = data_ini
    i = 0
    while d <= data_fim:
        recs = download_day(d)
        if recs:
            all_records.extend(recs)
        else:
            falhas.append(str(d))
        i += 1
        progress.progress(i / n_dias, text=f"A descarregar {d.strftime('%d/%m/%Y')} ({i}/{n_dias})")
        d += dt.timedelta(days=1)
        time.sleep(0.05)

    progress.empty()

    if not all_records:
        st.error("Nenhum dado descarregado. Verifica a ligacao a internet ou tenta outro periodo.")
        st.stop()

    n_ok = len(set(r["Data"] for r in all_records))
    status.success(f"{len(all_records)} registos horarios | {n_ok} dias com dados")

    if falhas:
        st.warning(f"Dias sem dados ({len(falhas)}): {', '.join(falhas[:10])}" +
                   (" ..." if len(falhas) > 10 else ""))

    # Metricas
    todos = [r["Preco_EUR_MWh"] for r in all_records]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Media (EUR/MWh)",  f"{statistics.mean(todos):.2f}")
    m2.metric("Min (EUR/MWh)",    f"{min(todos):.2f}")
    m3.metric("Max (EUR/MWh)",    f"{max(todos):.2f}")
    m4.metric("Horas negativas",  sum(1 for p in todos if p < 0))

    label = f"{data_ini.strftime('%d-%m-%Y')} a {data_fim.strftime('%d-%m-%Y')}"

    # --- CSV 1: Historico horario ---
    buf1 = io.StringIO()
    writer = csv.DictWriter(buf1, fieldnames=all_records[0].keys(), delimiter=";")
    writer.writeheader()
    writer.writerows(all_records)

    # --- CSV 2: Resumo mensal ---
    monthly = defaultdict(list)
    for r in all_records:
        monthly[r["Mes"]].append(r["Preco_EUR_MWh"])

    buf2 = io.StringIO()
    w2 = csv.writer(buf2, delimiter=";")
    w2.writerow(["Mes","N_Horas","Min_EUR_MWh","Max_EUR_MWh","Media_EUR_MWh",
                 "Mediana_EUR_MWh","Desvio_Padrao","Horas_Negativas","Pct_Negativas",
                 "Horas_Acima_100","Pct_Acima_100","Spread_Max_Min"])
    for m in sorted(monthly):
        prices = monthly[m]
        n   = len(prices)
        neg = sum(1 for p in prices if p < 0)
        hi  = sum(1 for p in prices if p > 100)
        w2.writerow([m, n,
                     round(min(prices),2), round(max(prices),2),
                     round(statistics.mean(prices),2),
                     round(statistics.median(prices),2),
                     round(statistics.stdev(prices) if n > 1 else 0, 2),
                     neg, round(neg/n*100, 1),
                     hi,  round(hi/n*100, 1),
                     round(max(prices)-min(prices), 2)])

    # --- CSV 3: Arbitragem BESS ---
    day_recs = defaultdict(list)
    for rec in all_records:
        day_recs[rec["Data"]].append(rec)

    arb_rows = []
    for d_str in sorted(day_recs):
        recs   = day_recs[d_str]
        min_r  = min(recs, key=lambda x: x["Preco_EUR_MWh"])
        max_r  = max(recs, key=lambda x: x["Preco_EUR_MWh"])
        spread = max_r["Preco_EUR_MWh"] - min_r["Preco_EUR_MWh"]
        rev    = spread * potencia * eficiencia
        arb_rows.append([d_str,
                         min_r["Preco_EUR_MWh"], min_r["Hora"],
                         max_r["Preco_EUR_MWh"], max_r["Hora"],
                         round(spread, 2), round(rev, 2)])

    buf3 = io.StringIO()
    w3 = csv.writer(buf3, delimiter=";")
    w3.writerow(["Data","Min_EUR_MWh","Hora_Carga","Max_EUR_MWh","Hora_Descarga",
                 "Spread_EUR_MWh","Receita_Bruta_EUR"])
    w3.writerows(arb_rows)

    # Resumo financeiro
    total_rev  = sum(r[6] for r in arb_rows)
    avg_spread = statistics.mean(r[5] for r in arb_rows)
    total_opex = opex * capacidade
    capex_tot  = capex * capacidade * 1000
    annuity    = capex_tot * (wacc*(1+wacc)**vida_util) / ((1+wacc)**vida_util - 1)
    ebitda     = total_rev - total_opex
    cashflow   = ebitda - annuity

    st.divider()
    st.subheader("Resumo Financeiro BESS")
    fa, fb = st.columns(2)
    fa.metric("Receita Bruta (EUR)", f"{total_rev:,.0f}")
    fb.metric("EBITDA (EUR)",        f"{ebitda:,.0f}")
    fc, fd = st.columns(2)
    fc.metric("Anuidade CAPEX (EUR/ano)", f"{annuity:,.0f}")
    fd.metric("Cash Flow Liquido (EUR/ano)", f"{cashflow:,.0f}")
    fe, ff = st.columns(2)
    fe.metric("Spread Medio Diario (EUR/MWh)", f"{avg_spread:.2f}")
    ff.metric("Dias com Spread > 50 EUR/MWh",  sum(1 for r in arb_rows if r[5] > 50))

    st.divider()
    st.subheader("Descarregar dados")

    nome_base = f"OMIE_Portugal_{data_ini.strftime('%d%m%Y')}_{data_fim.strftime('%d%m%Y')}"

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.download_button(
            label="Historico Horario (CSV)",
            data=buf1.getvalue().encode("utf-8"),
            file_name=f"{nome_base}_Historico.csv",
            mime="text/csv",
            use_container_width=True,
            type="primary",
        )
    with col_b:
        st.download_button(
            label="Resumo Mensal (CSV)",
            data=buf2.getvalue().encode("utf-8"),
            file_name=f"{nome_base}_Mensal.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col_c:
        st.download_button(
            label="Arbitragem BESS (CSV)",
            data=buf3.getvalue().encode("utf-8"),
            file_name=f"{nome_base}_BESS.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.caption("Ficheiros CSV separados por ponto e virgula (;) - abre directamente no Excel")
