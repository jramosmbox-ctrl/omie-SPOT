import io
import time
import statistics
import datetime as dt
from collections import defaultdict

import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

st.set_page_config(
    page_title="OMIE Portugal - Precos Horarios",
    page_icon="energy",
    layout="centered",
)

st.title("OMIE Portugal - Precos Horarios + Analise BESS")
st.caption("Dados reais do mercado diario MIBEL - Fonte: www.omie.es")
st.divider()

# --- Formulario de datas ---
st.subheader("Periodo de dados")

col1, col2 = st.columns(2)
with col1:
    data_ini = st.date_input(
        "Data de inicio",
        value=dt.date(2025, 1, 1),
        min_value=dt.date(2015, 1, 1),
        max_value=dt.date.today(),
        format="DD/MM/YYYY",
    )
with col2:
    data_fim = st.date_input(
        "Data de fim",
        value=dt.date(2025, 3, 31),
        min_value=dt.date(2015, 1, 1),
        max_value=dt.date.today(),
        format="DD/MM/YYYY",
    )

if data_fim < data_ini:
    st.error("A data de fim tem de ser posterior a data de inicio.")
    st.stop()

n_dias = (data_fim - data_ini).days + 1
tempo_est = max(1, round(n_dias * 0.4 / 60, 1))
st.info(f"**{n_dias} dias** selecionados - Download estimado: ~{tempo_est} min")

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
    capex     = st.number_input("CAPEX (EUR/kWh)",   value=250, min_value=0, step=10)
    vida_util = st.number_input("Vida util (anos)",  value=15,  min_value=1, max_value=30)
    wacc      = st.number_input("WACC (%)",          value=7,   min_value=1, max_value=30) / 100

st.divider()

# --- Download OMIE ---
OMIE_URL = (
    "https://www.omie.es/sites/default/files/dados/"
    "AGNO_{year}/MES_{month:02d}/TXT/"
    "marginalpdbcpt_{year}_{month:02d}_{day:02d}.1"
)
HEADERS = {"User-Agent": "Mozilla/5.0"}

def download_day(d):
    url = OMIE_URL.format(year=d.year, month=d.month, day=d.day)
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
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
            h0 = hour - 1
            wday = d.weekday()
            records.append({
                "Data"      : str(d),
                "Hora"      : hour,
                "Hora_Label": f"{h0:02d}:00-{hour:02d}:00",
                "Timestamp" : f"{str(d)} {h0:02d}:00",
                "Preco"     : round(price, 2),
                "Mes"       : str(d)[:7],
                "DiaSemana" : ["Seg","Ter","Qua","Qui","Sex","Sab","Dom"][wday % 7],
                "Periodo"   : (
                    "Vazio" if h0 in range(0, 7) else
                    "Cheia" if h0 in range(7, 10) or h0 in range(20, 24) else
                    "Ponta"
                ),
            })
        return records if records else None
    except Exception:
        return None

# --- Botao principal ---
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
        pct = i / n_dias
        progress.progress(pct, text=f"A descarregar {d.strftime('%d/%m/%Y')}... ({i}/{n_dias})")
        d += dt.timedelta(days=1)
        time.sleep(0.05)

    progress.empty()

    if not all_records:
        st.error("Nenhum dado descarregado. O OMIE pode estar indisponivel ou o URL mudou.")
        st.stop()

    n_dias_ok = len(set(r["Data"] for r in all_records))
    status.success(f"{len(all_records)} registos horarios descarregados ({n_dias_ok} dias ok)")

    if falhas:
        st.warning(f"Dias sem dados ({len(falhas)}): {', '.join(falhas[:10])}" +
                   (" ..." if len(falhas) > 10 else ""))

    todos = [r["Preco"] for r in all_records]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Media (EUR/MWh)",  f"{statistics.mean(todos):.2f}")
    m2.metric("Min (EUR/MWh)",    f"{min(todos):.2f}")
    m3.metric("Max (EUR/MWh)",    f"{max(todos):.2f}")
    m4.metric("Horas negativas",  sum(1 for p in todos if p < 0))

    # --- Excel ---
    BLUE_DARK  = "1F3864"
    BLUE_MID   = "2E75B6"
    WHITE      = "FFFFFF"
    GREY_LIGHT = "F2F2F2"

    def hfill(h):
        return PatternFill("solid", start_color=h, fgColor=h)

    def bdr():
        s = Side(style="thin", color="BFBFBF")
        return Border(left=s, right=s, top=s, bottom=s)

    def hcell(ws, row, col, val, bg=BLUE_MID, bold=True, size=10,
               color=WHITE, wrap=False, align="center"):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(name="Arial", bold=bold, size=size, color=color)
        c.fill      = hfill(bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        c.border    = bdr()
        return c

    def title_block(ws, merge, text, bg=BLUE_DARK, size=13):
        ws.merge_cells(merge)
        c = ws[merge.split(":")[0]]
        c.value     = text
        c.font      = Font(name="Arial", bold=True, size=size, color=WHITE)
        c.fill      = hfill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[int("".join(filter(str.isdigit, merge.split(":")[0])))].height = 28

    label = f"{data_ini.strftime('%d-%m-%Y')} a {data_fim.strftime('%d-%m-%Y')}"

    with st.spinner("A gerar ficheiro Excel..."):
        wb = Workbook()

        # Sheet 1
        ws1 = wb.active
        ws1.title = "Historico Horario"
        title_block(ws1, "A1:I1", f"OMIE Portugal - Precos Horarios | {label} (EUR/MWh)", size=13)
        ws1.merge_cells("A2:I2")
        ws1["A2"].value = f"Fonte: OMIE - www.omie.es | {len(all_records)} registos horarios reais"
        ws1["A2"].font = Font(name="Arial", italic=True, size=9, color="595959")
        ws1["A2"].alignment = Alignment(horizontal="center")
        cols1 = [("Data",12),("Hora",7),("Periodo Horario",16),("Timestamp",20),
                 ("Preco (EUR/MWh)",16),("Mes",10),("Dia Semana",12),
                 ("Periodo Tarifario",18),("Ano",7)]
        for i,(h_,w) in enumerate(cols1, 1):
            hcell(ws1, 3, i, h_, wrap=True)
            ws1.column_dimensions[get_column_letter(i)].width = w
        ws1.row_dimensions[3].height = 30
        ws1.freeze_panes = "A4"
        for r_i, rec in enumerate(all_records, 4):
            fill = hfill(GREY_LIGHT) if r_i % 2 == 0 else hfill(WHITE)
            row_data = [rec["Data"], rec["Hora"], rec["Hora_Label"], rec["Timestamp"],
                        rec["Preco"], rec["Mes"], rec["DiaSemana"], rec["Periodo"],
                        int(rec["Data"][:4])]
            for c_i, val in enumerate(row_data, 1):
                cell = ws1.cell(row=r_i, column=c_i, value=val)
                cell.font = Font(name="Arial", size=9)
                cell.fill = fill; cell.border = bdr()
                cell.alignment = Alignment(horizontal="center")
                if c_i == 5:
                    cell.number_format = "#,##0.00"
                    if isinstance(val, float):
                        if val < 0:
                            cell.font = Font(name="Arial", size=9, color="FF0000", bold=True)
                        elif val > 100:
                            cell.font = Font(name="Arial", size=9, color="7030A0", bold=True)
                if c_i in (1, 6):
                    cell.alignment = Alignment(horizontal="left")
        ws1.conditional_formatting.add(f"E4:E{3+len(all_records)}", ColorScaleRule(
            start_type="min",  start_color="63BE7B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max",    end_color="F8696B",
        ))

        # Sheet 2
        ws2 = wb.create_sheet("Resumo Mensal")
        title_block(ws2, "A1:L1", f"OMIE Portugal - Resumo Mensal (EUR/MWh) | {label}")
        monthly = defaultdict(list)
        for r in all_records: monthly[r["Mes"]].append(r["Preco"])
        mnames = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
        month_pt = {f"{y}-{m:02d}": mnames[m-1]+f"/{y}"
                    for y in range(data_ini.year, data_fim.year+1) for m in range(1,13)}
        hdrs2 = [("Mes",12),("N Horas",10),("Min (EUR/MWh)",13),("Max (EUR/MWh)",13),
                 ("Media (EUR/MWh)",14),("Mediana (EUR/MWh)",15),("Desvio Padrao",14),
                 ("Horas Negativas",16),("% Horas Neg.",14),
                 ("Horas > 100 EUR",13),("% Horas > 100",15),("Spread Max-Min",15)]
        for i,(h_,w) in enumerate(hdrs2, 1):
            hcell(ws2, 2, i, h_, wrap=True)
            ws2.column_dimensions[get_column_letter(i)].width = w
        ws2.row_dimensions[2].height = 36
        all_p = []
        for r_i, m in enumerate(sorted(monthly), 3):
            prices = monthly[m]; all_p.extend(prices)
            n = len(prices)
            neg = sum(1 for p in prices if p < 0)
            hi  = sum(1 for p in prices if p > 100)
            row = [month_pt.get(m,m), n,
                   round(min(prices),2), round(max(prices),2),
                   round(statistics.mean(prices),2), round(statistics.median(prices),2),
                   round(statistics.stdev(prices) if n>1 else 0,2),
                   neg, neg/n, hi, hi/n, round(max(prices)-min(prices),2)]
            fill = hfill(GREY_LIGHT) if r_i%2==0 else hfill(WHITE)
            for c_i,val in enumerate(row,1):
                cell = ws2.cell(row=r_i,column=c_i,value=val)
                cell.font=Font(name="Arial",size=10); cell.fill=fill
                cell.border=bdr(); cell.alignment=Alignment(horizontal="center")
                if c_i in (3,4,5,6,7,12): cell.number_format="#,##0.00"
                if c_i in (9,11):         cell.number_format="0.0%"
        tr=3+len(monthly); n_a=len(all_p)
        neg_a=sum(1 for p in all_p if p<0); hi_a=sum(1 for p in all_p if p>100)
        totals=["TOTAL",n_a,round(min(all_p),2),round(max(all_p),2),
                round(statistics.mean(all_p),2),round(statistics.median(all_p),2),
                round(statistics.stdev(all_p),2),neg_a,neg_a/n_a,hi_a,hi_a/n_a,
                round(max(all_p)-min(all_p),2)]
        for c_i,val in enumerate(totals,1):
            cell=ws2.cell(row=tr,column=c_i,value=val)
            cell.font=Font(name="Arial",bold=True,size=10,color=WHITE)
            cell.fill=hfill(BLUE_DARK); cell.border=bdr()
            cell.alignment=Alignment(horizontal="center")
            if c_i in (3,4,5,6,7,12): cell.number_format="#,##0.00"
            if c_i in (9,11):         cell.number_format="0.0%"
        ws2.freeze_panes="A3"

        # Sheet 3
        ws3 = wb.create_sheet("Perfil Horario")
        title_block(ws3,"A1:G1",f"OMIE Portugal - Perfil Horario Medio (EUR/MWh) | {label}")
        hourly=defaultdict(list)
        for r in all_records: hourly[r["Hora"]].append(r["Preco"])
        hdrs3=[("Hora",8),("Periodo",10),("Media (EUR/MWh)",14),("Min (EUR/MWh)",13),
               ("Max (EUR/MWh)",13),("Mediana (EUR/MWh)",15),("Desvio Padrao",14)]
        for i,(h_,w) in enumerate(hdrs3,1):
            hcell(ws3,2,i,h_,wrap=True)
            ws3.column_dimensions[get_column_letter(i)].width=w
        ws3.row_dimensions[2].height=30
        PC={"Vazio":"EBF3FB","Cheia":"FFF2CC","Ponta":"FCE4D6"}
        def periodo_h(h):
            h0=h-1
            if h0 in range(0,7): return "Vazio"
            if h0 in range(7,10) or h0 in range(20,24): return "Cheia"
            return "Ponta"
        for h in range(1,25):
            prices=hourly.get(h,[0]); p=periodo_h(h)
            row=[f"{(h-1):02d}:00",p,round(statistics.mean(prices),2),round(min(prices),2),
                 round(max(prices),2),round(statistics.median(prices),2),
                 round(statistics.stdev(prices) if len(prices)>1 else 0,2)]
            for c_i,val in enumerate(row,1):
                cell=ws3.cell(row=h+2,column=c_i,value=val)
                cell.font=Font(name="Arial",size=10); cell.fill=hfill(PC[p])
                cell.border=bdr(); cell.alignment=Alignment(horizontal="center")
                if c_i in (3,4,5,6,7): cell.number_format="#,##0.00"

        # Sheet 4
        ws4=wb.create_sheet("Analise BESS")
        title_block(ws4,"A1:K1",f"Analise de Arbitragem BESS - Portugal | {label}",size=13)
        ws4.merge_cells("A3:D3"); hcell(ws4,3,1,"PARAMETROS DO SISTEMA BESS",size=11)
        for c in range(2,5): ws4.cell(row=3,column=c).fill=hfill(BLUE_MID)
        params_list=[("Capacidade (MWh)",capacidade,"MWh"),("Potencia (MW)",potencia,"MW"),
                     ("Eficiencia RT",eficiencia,"%"),("OPEX (EUR/MWh/ano)",opex,"EUR/MWh"),
                     ("CAPEX (EUR/kWh)",capex,"EUR/kWh"),("Vida util (anos)",vida_util,"anos"),
                     ("WACC",wacc,"%")]
        for i,lbl in enumerate(["Parametro","Valor","Unidade"],1):
            hcell(ws4,4,i,lbl,bg="4472C4",size=9)
        for r_i,(name,val,unit) in enumerate(params_list,5):
            fill=hfill(GREY_LIGHT) if r_i%2==0 else hfill(WHITE)
            for c_i,v in enumerate([name,val,unit],1):
                cell=ws4.cell(row=r_i,column=c_i,value=v)
                cell.font=Font(name="Arial",size=9,color="0000FF" if c_i==2 else "000000",bold=(c_i==2))
                cell.fill=fill; cell.border=bdr(); cell.alignment=Alignment(horizontal="center")
        ws4.column_dimensions["A"].width=30; ws4.column_dimensions["B"].width=12; ws4.column_dimensions["C"].width=14
        ws4.merge_cells("E3:K3"); hcell(ws4,3,5,"ARBITRAGEM DIARIA - 1 CICLO/DIA",size=11)
        for c in range(6,12): ws4.cell(row=3,column=c).fill=hfill(BLUE_MID)
        arb_hdrs=[("Data",12),("Min (EUR/MWh)",13),("H. Carga",9),("Max (EUR/MWh)",13),
                  ("H. Descarga",11),("Spread (EUR/MWh)",14),("Receita Bruta (EUR)",16)]
        for i,(h_,w) in enumerate(arb_hdrs,5):
            hcell(ws4,4,i,h_,bg="4472C4",size=9,wrap=True)
            ws4.column_dimensions[get_column_letter(i)].width=w
        ws4.row_dimensions[4].height=30
        day_recs=defaultdict(list)
        for rec in all_records: day_recs[rec["Data"]].append(rec)
        arb_rows=[]
        for d_str in sorted(day_recs):
            recs=day_recs[d_str]
            min_r=min(recs,key=lambda x:x["Preco"]); max_r=max(recs,key=lambda x:x["Preco"])
            spread=max_r["Preco"]-min_r["Preco"]; rev=spread*potencia*eficiencia
            arb_rows.append((d_str,min_r["Preco"],min_r["Hora"],max_r["Preco"],max_r["Hora"],spread,rev))
        for r_i,row in enumerate(arb_rows,5):
            fill=hfill(GREY_LIGHT) if r_i%2==0 else hfill(WHITE)
            for c_i,val in enumerate(row,5):
                cell=ws4.cell(row=r_i,column=c_i,value=val)
                cell.font=Font(name="Arial",size=9); cell.fill=fill
                cell.border=bdr(); cell.alignment=Alignment(horizontal="center")
                if c_i in (6,7,8,11): cell.number_format="#,##0.00"
                if c_i==10 and isinstance(val,float) and val>50:
                    cell.font=Font(name="Arial",size=9,color="00B050",bold=True)
        total_rev=sum(r[6] for r in arb_rows); avg_spread=statistics.mean(r[5] for r in arb_rows)
        total_opex=opex*capacidade; capex_tot=capex*capacidade*1000
        annuity=capex_tot*(wacc*(1+wacc)**vida_util)/((1+wacc)**vida_util-1)
        ebitda=total_rev-total_opex
        sr=5+len(params_list)+3
        ws4.merge_cells(f"A{sr}:D{sr}"); hcell(ws4,sr,1,"RESUMO FINANCEIRO",size=11)
        for c in range(2,5): ws4.cell(row=sr,column=c).fill=hfill(BLUE_MID)
        summary=[("Receita Bruta Total (EUR)",total_rev,"EUR"),
                 ("OPEX Anual (EUR)",total_opex,"EUR"),("EBITDA (EUR)",ebitda,"EUR"),
                 ("CAPEX Total (EUR)",capex_tot,"EUR"),("Anuidade CAPEX (EUR/ano)",annuity,"EUR/ano"),
                 ("Cash Flow Liquido Anual (EUR)",ebitda-annuity,"EUR"),
                 ("Spread Medio Diario (EUR/MWh)",avg_spread,"EUR/MWh"),
                 ("Dias com Spread > 50 EUR/MWh",sum(1 for r in arb_rows if r[5]>50),"dias"),
                 ("Dias com Spread > 100 EUR/MWh",sum(1 for r in arb_rows if r[5]>100),"dias")]
        for i,(lbl,val,unit) in enumerate(summary,sr+1):
            fill=hfill(GREY_LIGHT) if i%2==0 else hfill(WHITE)
            ws4.cell(row=i,column=1,value=lbl).font=Font(name="Arial",size=9)
            ws4.cell(row=i,column=1).fill=fill; ws4.cell(row=i,column=1).border=bdr()
            cell_v=ws4.cell(row=i,column=2,value=round(val,2))
            cell_v.font=Font(name="Arial",size=9,bold=True); cell_v.number_format="#,##0.00"
            cell_v.fill=fill; cell_v.border=bdr(); cell_v.alignment=Alignment(horizontal="center")
            ws4.cell(row=i,column=3,value=unit).font=Font(size=9)
            ws4.cell(row=i,column=3).fill=fill; ws4.cell(row=i,column=3).border=bdr()

        buf=io.BytesIO(); wb.save(buf); buf.seek(0)

    nome=(f"OMIE_Portugal_{data_ini.strftime('%d%m%Y')}_{data_fim.strftime('%d%m%Y')}_Precos_BESS.xlsx")
    st.download_button(
        label="Descarregar Excel",
        data=buf, file_name=nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )
    st.caption(f"Ficheiro: {nome} | 4 folhas | {len(all_records)} registos horarios")
