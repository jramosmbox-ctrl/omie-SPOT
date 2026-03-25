import io
import time
import statistics
import datetime as dt
from collections import defaultdict

import requests
import xlsxwriter
import streamlit as st

st.set_page_config(page_title="OMIE Portugal - Precos Horarios", layout="centered")

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
                "Preco"    : round(price, 2),
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
if st.button("Descarregar dados e gerar Excel", type="primary", use_container_width=True):

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

    todos = [r["Preco"] for r in all_records]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Media (EUR/MWh)",  f"{statistics.mean(todos):.2f}")
    m2.metric("Min (EUR/MWh)",    f"{min(todos):.2f}")
    m3.metric("Max (EUR/MWh)",    f"{max(todos):.2f}")
    m4.metric("Horas negativas",  sum(1 for p in todos if p < 0))

    label = f"{data_ini.strftime('%d-%m-%Y')} a {data_fim.strftime('%d-%m-%Y')}"

    with st.spinner("A gerar ficheiro Excel..."):

        buf = io.BytesIO()
        wb  = xlsxwriter.Workbook(buf, {"in_memory": True})

        # --- Formatos ---
        hdr_dark  = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1F3864",
                                   "align":"center","valign":"vcenter","border":1,"font_name":"Arial","font_size":11})
        hdr_mid   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E75B6",
                                   "align":"center","valign":"vcenter","border":1,"font_name":"Arial","font_size":10,"text_wrap":True})
        hdr_blue2 = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#4472C4",
                                   "align":"center","valign":"vcenter","border":1,"font_name":"Arial","font_size":9,"text_wrap":True})
        cell_norm = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"center"})
        cell_odd  = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"center","bg_color":"#F2F2F2"})
        cell_left = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"left"})
        cell_odd_l= wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"left","bg_color":"#F2F2F2"})
        num_fmt   = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"center","num_format":"#,##0.00"})
        num_odd   = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"center","num_format":"#,##0.00","bg_color":"#F2F2F2"})
        num_neg   = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"center","num_format":"#,##0.00","font_color":"#FF0000","bold":True})
        num_high  = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"center","num_format":"#,##0.00","font_color":"#7030A0","bold":True})
        num_green = wb.add_format({"font_name":"Arial","font_size":9,"border":1,"align":"center","num_format":"#,##0.00","font_color":"#00B050","bold":True})
        pct_fmt   = wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center","num_format":"0.0%"})
        pct_odd   = wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center","num_format":"0.0%","bg_color":"#F2F2F2"})
        num10     = wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center","num_format":"#,##0.00"})
        num10_odd = wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center","num_format":"#,##0.00","bg_color":"#F2F2F2"})
        cell10    = wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center"})
        cell10_odd= wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center","bg_color":"#F2F2F2"})
        tot_fmt   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1F3864",
                                   "align":"center","border":1,"font_name":"Arial","font_size":10,"num_format":"#,##0.00"})
        tot_pct   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1F3864",
                                   "align":"center","border":1,"font_name":"Arial","font_size":10,"num_format":"0.0%"})
        tot_int   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1F3864",
                                   "align":"center","border":1,"font_name":"Arial","font_size":10})
        param_val = wb.add_format({"bold":True,"font_color":"#0000FF","font_name":"Arial",
                                   "font_size":9,"border":1,"align":"center"})
        param_val_odd = wb.add_format({"bold":True,"font_color":"#0000FF","font_name":"Arial",
                                       "font_size":9,"border":1,"align":"center","bg_color":"#F2F2F2"})

        # ================================================================
        # SHEET 1 - Historico Horario
        # ================================================================
        ws1 = wb.add_worksheet("Historico Horario")
        ws1.set_row(0, 28)
        ws1.set_row(2, 30)
        ws1.freeze_panes(3, 0)

        ws1.merge_range("A1:I1", f"OMIE Portugal - Precos Horarios | {label} (EUR/MWh)", hdr_dark)
        ws1.merge_range("A2:I2", f"Fonte: OMIE - www.omie.es | {len(all_records)} registos horarios reais",
                        wb.add_format({"italic":True,"font_size":9,"font_color":"#595959","align":"center","font_name":"Arial"}))

        hdrs1 = ["Data","Hora","Periodo Horario","Timestamp","Preco (EUR/MWh)",
                 "Mes","Dia Semana","Periodo Tarifario","Ano"]
        widths1 = [12,7,16,20,16,10,12,18,7]
        for i,(h,w) in enumerate(zip(hdrs1,widths1)):
            ws1.write(2, i, h, hdr_mid)
            ws1.set_column(i, i, w)

        for r_i, rec in enumerate(all_records):
            row = r_i + 3
            odd = (row % 2 == 0)
            cf  = cell_odd if odd else cell_norm
            nf  = num_odd  if odd else num_fmt
            lf  = cell_odd_l if odd else cell_left
            price = rec["Preco"]
            pf = num_neg if price < 0 else (num_high if price > 100 else nf)
            ws1.write(row, 0, rec["Data"],      lf)
            ws1.write(row, 1, rec["Hora"],      cf)
            ws1.write(row, 2, rec["Horario"],   cf)
            ws1.write(row, 3, rec["Timestamp"], cf)
            ws1.write(row, 4, price,            pf)
            ws1.write(row, 5, rec["Mes"],       lf)
            ws1.write(row, 6, rec["DiaSemana"], cf)
            ws1.write(row, 7, rec["Periodo"],   cf)
            ws1.write(row, 8, int(rec["Data"][:4]), cf)

        # ================================================================
        # SHEET 2 - Resumo Mensal
        # ================================================================
        ws2 = wb.add_worksheet("Resumo Mensal")
        ws2.set_row(0, 28)
        ws2.set_row(1, 36)
        ws2.freeze_panes(2, 0)

        ws2.merge_range("A1:L1", f"OMIE Portugal - Resumo Mensal (EUR/MWh) | {label}", hdr_dark)

        hdrs2 = ["Mes","N Horas","Min (EUR/MWh)","Max (EUR/MWh)","Media (EUR/MWh)",
                 "Mediana (EUR/MWh)","Desvio Padrao","Horas Negativas","% Neg.",
                 "Horas > 100 EUR","% > 100 EUR","Spread Max-Min"]
        widths2 = [12,10,13,13,14,15,14,16,10,13,11,14]
        for i,(h,w) in enumerate(zip(hdrs2,widths2)):
            ws2.write(1, i, h, hdr_mid)
            ws2.set_column(i, i, w)

        monthly = defaultdict(list)
        for r in all_records: monthly[r["Mes"]].append(r["Preco"])
        mnames = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
        month_pt = {f"{y}-{m:02d}": mnames[m-1]+f"/{y}"
                    for y in range(data_ini.year, data_fim.year+1) for m in range(1,13)}

        all_p = []
        for r_i, m in enumerate(sorted(monthly)):
            prices = monthly[m]; all_p.extend(prices)
            n   = len(prices)
            neg = sum(1 for p in prices if p < 0)
            hi  = sum(1 for p in prices if p > 100)
            row = r_i + 2
            odd = (row % 2 == 0)
            nf  = num10_odd if odd else num10
            cf  = cell10_odd if odd else cell10
            pf  = pct_odd   if odd else pct_fmt
            ws2.write(row, 0,  month_pt.get(m,m),                        cf)
            ws2.write(row, 1,  n,                                         cf)
            ws2.write(row, 2,  round(min(prices),2),                      nf)
            ws2.write(row, 3,  round(max(prices),2),                      nf)
            ws2.write(row, 4,  round(statistics.mean(prices),2),          nf)
            ws2.write(row, 5,  round(statistics.median(prices),2),        nf)
            ws2.write(row, 6,  round(statistics.stdev(prices) if n>1 else 0,2), nf)
            ws2.write(row, 7,  neg,                                       cf)
            ws2.write(row, 8,  neg/n,                                     pf)
            ws2.write(row, 9,  hi,                                        cf)
            ws2.write(row, 10, hi/n,                                      pf)
            ws2.write(row, 11, round(max(prices)-min(prices),2),          nf)

        tr  = len(monthly) + 2
        n_a = len(all_p); neg_a = sum(1 for p in all_p if p<0); hi_a = sum(1 for p in all_p if p>100)
        ws2.write(tr, 0,  "TOTAL",                                     tot_int)
        ws2.write(tr, 1,  n_a,                                         tot_int)
        ws2.write(tr, 2,  round(min(all_p),2),                         tot_fmt)
        ws2.write(tr, 3,  round(max(all_p),2),                         tot_fmt)
        ws2.write(tr, 4,  round(statistics.mean(all_p),2),             tot_fmt)
        ws2.write(tr, 5,  round(statistics.median(all_p),2),           tot_fmt)
        ws2.write(tr, 6,  round(statistics.stdev(all_p),2),            tot_fmt)
        ws2.write(tr, 7,  neg_a,                                        tot_int)
        ws2.write(tr, 8,  neg_a/n_a,                                    tot_pct)
        ws2.write(tr, 9,  hi_a,                                         tot_int)
        ws2.write(tr, 10, hi_a/n_a,                                     tot_pct)
        ws2.write(tr, 11, round(max(all_p)-min(all_p),2),              tot_fmt)

        # ================================================================
        # SHEET 3 - Perfil Horario
        # ================================================================
        ws3 = wb.add_worksheet("Perfil Horario")
        ws3.set_row(0, 28)
        ws3.set_row(1, 30)

        ws3.merge_range("A1:G1", f"OMIE Portugal - Perfil Horario Medio (EUR/MWh) | {label}", hdr_dark)

        hdrs3 = ["Hora","Periodo","Media (EUR/MWh)","Min (EUR/MWh)","Max (EUR/MWh)",
                 "Mediana (EUR/MWh)","Desvio Padrao"]
        widths3 = [8,10,14,13,13,15,14]
        for i,(h,w) in enumerate(zip(hdrs3,widths3)):
            ws3.write(1, i, h, hdr_mid)
            ws3.set_column(i, i, w)

        hourly = defaultdict(list)
        for r in all_records: hourly[r["Hora"]].append(r["Preco"])

        PC_BG = {"Vazio":"#EBF3FB","Cheia":"#FFF2CC","Ponta":"#FCE4D6"}
        def periodo_h(h):
            h0 = h-1
            if h0 in range(0,7): return "Vazio"
            if h0 in range(7,10) or h0 in range(20,24): return "Cheia"
            return "Ponta"

        for h in range(1,25):
            prices = hourly.get(h,[0]); p = periodo_h(h)
            bg = PC_BG[p]
            cf_h = wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center","bg_color":bg})
            nf_h = wb.add_format({"font_name":"Arial","font_size":10,"border":1,"align":"center","bg_color":bg,"num_format":"#,##0.00"})
            row = h + 1
            ws3.write(row, 0, f"{(h-1):02d}:00",                                              cf_h)
            ws3.write(row, 1, p,                                                               cf_h)
            ws3.write(row, 2, round(statistics.mean(prices),2),                               nf_h)
            ws3.write(row, 3, round(min(prices),2),                                           nf_h)
            ws3.write(row, 4, round(max(prices),2),                                           nf_h)
            ws3.write(row, 5, round(statistics.median(prices),2),                             nf_h)
            ws3.write(row, 6, round(statistics.stdev(prices) if len(prices)>1 else 0,2),     nf_h)

        # ================================================================
        # SHEET 4 - Analise BESS
        # ================================================================
        ws4 = wb.add_worksheet("Analise BESS")
        ws4.set_row(0, 28)
        ws4.set_row(2, 28)
        ws4.set_row(3, 30)
        ws4.set_column(0, 0, 30)
        ws4.set_column(1, 1, 12)
        ws4.set_column(2, 2, 14)

        ws4.merge_range("A1:K1", f"Analise de Arbitragem BESS - Portugal | {label}", hdr_dark)
        ws4.merge_range("A3:D3", "PARAMETROS DO SISTEMA BESS", hdr_mid)
        ws4.write(3, 0, "Parametro", hdr_blue2)
        ws4.write(3, 1, "Valor",     hdr_blue2)
        ws4.write(3, 2, "Unidade",   hdr_blue2)

        params_list = [
            ("Capacidade (MWh)", capacidade, "MWh"),
            ("Potencia (MW)",    potencia,   "MW"),
            ("Eficiencia RT",    eficiencia, "%"),
            ("OPEX (EUR/MWh/ano)", opex,     "EUR/MWh"),
            ("CAPEX (EUR/kWh)",  capex,      "EUR/kWh"),
            ("Vida util (anos)", vida_util,  "anos"),
            ("WACC",             wacc,       "%"),
        ]
        for r_i,(name,val,unit) in enumerate(params_list):
            row = r_i + 4
            odd = (row % 2 == 0)
            cf  = cell_odd if odd else cell_norm
            pv  = param_val_odd if odd else param_val
            ws4.write(row, 0, name, cf)
            ws4.write(row, 1, val,  pv)
            ws4.write(row, 2, unit, cf)

        # Arbitragem diaria
        arb_col_start = 4
        ws4.set_column(4, 4, 12); ws4.set_column(5, 5, 13); ws4.set_column(6, 6, 9)
        ws4.set_column(7, 7, 13); ws4.set_column(8, 8, 11)
        ws4.set_column(9, 9, 14); ws4.set_column(10, 10, 16)

        ws4.merge_range("E3:K3", "ARBITRAGEM DIARIA - 1 CICLO/DIA", hdr_mid)
        arb_hdrs = ["Data","Min (EUR/MWh)","H. Carga","Max (EUR/MWh)","H. Descarga",
                    "Spread (EUR/MWh)","Receita Bruta (EUR)"]
        for i,h in enumerate(arb_hdrs):
            ws4.write(3, i+4, h, hdr_blue2)

        day_recs = defaultdict(list)
        for rec in all_records: day_recs[rec["Data"]].append(rec)

        arb_rows = []
        for d_str in sorted(day_recs):
            recs  = day_recs[d_str]
            min_r = min(recs, key=lambda x: x["Preco"])
            max_r = max(recs, key=lambda x: x["Preco"])
            spread = max_r["Preco"] - min_r["Preco"]
            rev    = spread * potencia * eficiencia
            arb_rows.append((d_str, min_r["Preco"], min_r["Hora"],
                             max_r["Preco"], max_r["Hora"], spread, rev))

        for r_i, row_data in enumerate(arb_rows):
            row = r_i + 4
            odd = (row % 2 == 0)
            cf  = cell_odd if odd else cell_norm
            nf  = num_odd  if odd else num_fmt
            spread_val = row_data[5]
            sf = num_green if spread_val > 50 else nf
            ws4.write(row, 4,  row_data[0], cf)
            ws4.write(row, 5,  row_data[1], nf)
            ws4.write(row, 6,  row_data[2], cf)
            ws4.write(row, 7,  row_data[3], nf)
            ws4.write(row, 8,  row_data[4], cf)
            ws4.write(row, 9,  spread_val,  sf)
            ws4.write(row, 10, row_data[6], nf)

        # Resumo financeiro
        total_rev  = sum(r[6] for r in arb_rows)
        avg_spread = statistics.mean(r[5] for r in arb_rows)
        total_opex = opex * capacidade
        capex_tot  = capex * capacidade * 1000
        annuity    = capex_tot * (wacc*(1+wacc)**vida_util) / ((1+wacc)**vida_util - 1)
        ebitda     = total_rev - total_opex
        cashflow   = ebitda - annuity

        sr = len(params_list) + 6
        ws4.merge_range(sr, 0, sr, 3, "RESUMO FINANCEIRO", hdr_mid)
        summary = [
            ("Receita Bruta Total (EUR)",     total_rev),
            ("OPEX Anual (EUR)",              total_opex),
            ("EBITDA (EUR)",                  ebitda),
            ("CAPEX Total (EUR)",             capex_tot),
            ("Anuidade CAPEX (EUR/ano)",       annuity),
            ("Cash Flow Liquido Anual (EUR)", cashflow),
            ("Spread Medio Diario (EUR/MWh)", avg_spread),
            ("Dias com Spread > 50 EUR/MWh",  sum(1 for r in arb_rows if r[5]>50)),
            ("Dias com Spread > 100 EUR/MWh", sum(1 for r in arb_rows if r[5]>100)),
        ]
        for i,(lbl,val) in enumerate(summary):
            row = sr + 1 + i
            odd = (row % 2 == 0)
            cf  = cell_odd if odd else cell_norm
            nf  = num_odd  if odd else num_fmt
            ws4.write(row, 0, lbl,             cf)
            ws4.write(row, 1, round(val, 2),   nf)

        wb.close()
        buf.seek(0)

    # Resumo financeiro na app
    st.divider()
    st.subheader("Resumo Financeiro BESS")
    fa, fb = st.columns(2)
    fa.metric("Receita Bruta (EUR)", f"{total_rev:,.0f}")
    fb.metric("EBITDA (EUR)",        f"{ebitda:,.0f}")
    fc, fd = st.columns(2)
    fc.metric("Anuidade CAPEX (EUR/ano)",    f"{annuity:,.0f}")
    fd.metric("Cash Flow Liquido (EUR/ano)", f"{cashflow:,.0f}")
    fe, ff = st.columns(2)
    fe.metric("Spread Medio Diario (EUR/MWh)", f"{avg_spread:.2f}")
    ff.metric("Dias com Spread > 50 EUR/MWh",  sum(1 for r in arb_rows if r[5] > 50))

    st.divider()
    nome = (f"OMIE_Portugal_{data_ini.strftime('%d%m%Y')}_"
            f"{data_fim.strftime('%d%m%Y')}_Precos_BESS.xlsx")
    st.download_button(
        label="Descarregar Excel (.xlsx)",
        data=buf,
        file_name=nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
    st.caption(f"Ficheiro: {nome} | 4 folhas | {len(all_records)} registos horarios")
