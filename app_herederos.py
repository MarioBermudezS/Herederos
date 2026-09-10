import io
import pandas as pd
import streamlit as st
from typing import Dict, List

st.set_page_config(layout="wide")
st.markdown("<style>.stDataFrame{font-size:10px!important}table{line-height:1!important}th,td{padding:1px!important}</style>", unsafe_allow_html=True)
st.title("Herederos")

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
CONCEPTOS = ["Ventas", "Coste Ventas", "MARGEN BRUTO", "R. B.", "Otros Ingresos", "Ingresos Operativos", "Gastos Personal", "Alquileres", "Reparaciones", "Seguros", "Suministros", "Otros Servicios", "TOTAL GASTOS OPERATIVOS", "Amortizaciones", "GASTOS ESTRUCTURA", "B.A.I.I.", "Gastos Financieros", "Ingresos Financieros", "RDO. FINANCIERO", "Resultados Extraordinarios", "B.A.I."]
NEGRITA = ["MARGEN BRUTO", "Ingresos Operativos", "TOTAL GASTOS OPERATIVOS", "GASTOS ESTRUCTURA", "B.A.I.I.", "RDO. FINANCIERO", "B.A.I."]
GASTOS_SET = {"Coste Ventas", "Gastos Personal", "Alquileres", "Reparaciones", "Seguros", "Suministros", "Otros Servicios", "TOTAL GASTOS OPERATIVOS", "Amortizaciones", "GASTOS ESTRUCTURA", "Gastos Financieros"}

@st.cache_data
def load_data():
    return pd.read_excel("BaseDatos2026.xlsx", sheet_name="BS")

def fmt_pct(v):
    return f"{v*100:,.1f}%".replace(",","X").replace(".",",").replace("X",".")

def fmt_eur(v):
    return f"{v:,.0f}€".replace(",","X").replace(".",",").replace("X",".")

def calcular_rb(df_f):
    if df_f.empty: return 0.0
    tm, tv = 0.0, 0.0
    for _, g in df_f.groupby(["Año", "Mes", "Departamento"]):
        v = g[g["Resultados"]=="Ventas"]["Importe D"].sum()
        rb = g[g["Resultados"].str.upper().str.strip().isin(["R. B.", "R.B.", "R.B"])]["Importe D"]
        rb_v = rb.iloc[0] if len(rb) > 0 else 0.0
        tv += v
        tm += v * rb_v
    return (tm/tv) if tv != 0 else 0.0

def calcular_res(df_f):
    if df_f.empty: return {c: 0.0 for c in CONCEPTOS}
    r = df_f.groupby("Resultados")["Importe D"].sum().to_dict()
    def gv(c): return r.get(c, 0.0)
    v = gv("Ventas")
    tm, tv = 0.0, 0.0
    for _, g in df_f.groupby(["Año", "Mes", "Departamento"]):
        vv = g[g["Resultados"]=="Ventas"]["Importe D"].sum()
        rb = g[g["Resultados"].str.upper().str.strip().isin(["R. B.", "R.B.", "R.B"])]["Importe D"]
        rb_v = rb.iloc[0] if len(rb) > 0 else 0.0
        tv += vv
        tm += vv * rb_v
    mb = tm
    rb = (mb/tv) if tv != 0 else 0.0
    if tv == 0 and v != 0: rb = gv("R. B."); mb = rb*v
    cv = v - mb
    oi = gv("Otros Ingresos")
    io = mb + oi
    gp = gv("Gastos Personal")
    gop = gv("Alquileres") + gv("Reparaciones") + gv("Seguros") + gv("Suministros") + gv("Otros Servicios")
    am = gv("Amortizaciones")
    ge = gop + gp + am
    baii = io - gp - gop - am
    gf = gv("Gastos Financieros")
    inf = gv("Ingresos Financieros")
    rf = gf + inf
    re = gv("Resultados Extraordinarios")
    bai = baii - rf - re
    return {"Ventas": v, "Coste Ventas": cv, "MARGEN BRUTO": mb, "R. B.": rb, "Otros Ingresos": oi, "Ingresos Operativos": io, "Gastos Personal": gp, "Alquileres": gv("Alquileres"), "Reparaciones": gv("Reparaciones"), "Seguros": gv("Seguros"), "Suministros": gv("Suministros"), "Otros Servicios": gv("Otros Servicios"), "TOTAL GASTOS OPERATIVOS": gop, "Amortizaciones": am, "GASTOS ESTRUCTURA": ge, "B.A.I.I.": baii, "Gastos Financieros": gf, "Ingresos Financieros": inf, "RDO. FINANCIERO": rf, "Resultados Extraordinarios": re, "B.A.I.": bai}

def renderizar(df_d, df_n=None, negrita=False, colores=False):
    def estilo_fila(row):
        if negrita:
            concepto = df_d.iloc[row.name, 0]
            if concepto in NEGRITA:
                return ['font-weight:bold;background:#f0f0f0']*len(row)
        return ['']*len(row)
    s = df_d.style.apply(estilo_fila, axis=1)
    if colores and df_n is not None:
        for col_idx, col in enumerate(df_n.columns[1:], 1):
            for row_idx in range(len(df_n)):
                concepto = df_d.iloc[row_idx, 0]
                vals = pd.to_numeric(df_n.iloc[1:, col_idx], errors='coerce')
                if vals.notna().any():
                    if concepto in GASTOS_SET:
                        min_idx = vals.idxmin()
                        if row_idx == min_idx:
                            s = s.applymap(lambda x: 'background:#90EE90', subset=pd.IndexSlice[row_idx, col])
                    else:
                        max_idx = vals.idxmax()
                        if row_idx == max_idx:
                            s = s.applymap(lambda x: 'background:#90EE90', subset=pd.IndexSlice[row_idx, col])
    st.dataframe(s, use_container_width=True, hide_index=True)

df = load_data()
st.sidebar.header("FILTROS")
modulo = st.sidebar.radio("Módulo", ["Resultados", "R.B.", "KPI"])
anos = st.sidebar.multiselect("Año(s)", sorted(df["Año"].unique()), default=[2026])
meses_sel = st.sidebar.multiselect("Mes(es)", [m for m in MESES if m in df["Mes"].unique()], default=["Enero"])
depts = sorted(df["Departamento"].dropna().unique())
tiendas = st.sidebar.multiselect("Tienda(s)", depts, default=depts[:1] if depts else [])

if not anos or not meses_sel or not tiendas:
    st.warning("Selecciona parámetros")
    st.stop()

df_f = df[(df["Año"].isin(anos)) & (df["Mes"].isin(meses_sel)) & (df["Departamento"].isin(tiendas))]

if modulo == "R.B.":
    st.subheader(f"R.B. {','.join(map(str, anos))}")
    cols = ["Tienda"] + meses_sel + (["Prom"] if len(meses_sel) > 1 else [])
    rows_d, rows_n = [], []
    for t in tiendas:
        fila_d, fila_n = [t], [t]
        for m in meses_sel:
            df_tm = df[(df["Año"].isin(anos)) & (df["Mes"]==m) & (df["Departamento"]==t)]
            v = calcular_rb(df_tm)
            fila_n.append(v)
            fila_d.append(fmt_pct(v))
        if len(meses_sel) > 1:
            df_ac = df[(df["Año"].isin(anos)) & (df["Mes"].isin(meses_sel)) & (df["Departamento"]==t)]
            v = calcular_rb(df_ac)
            fila_n.append(v)
            fila_d.append(fmt_pct(v))
        rows_d.append(fila_d)
        rows_n.append(fila_n)
    df_d, df_n = pd.DataFrame(rows_d, columns=cols), pd.DataFrame(rows_n, columns=cols)
    renderizar(df_d, df_n)

elif modulo == "KPI":
    st.subheader(f"KPI {','.join(map(str, anos))}")
    cols_tabla = ["Concepto"] + tiendas
    rows_d, rows_n = [], []
    for c in CONCEPTOS:
        fila_d, fila_n = [c], [c]
        for t in tiendas:
            df_t = df[(df["Año"].isin(anos)) & (df["Mes"].isin(meses_sel)) & (df["Departamento"]==t)]
            res_t = calcular_res(df_t)
            val = res_t.get(c, 0.0)
            v_v = res_t.get("Ventas", 1.0)
            kpi = val if c == "R. B." else (val/v_v if v_v != 0 else 0.0)
            fila_n.append(kpi)
            fila_d.append(fmt_pct(kpi))
        rows_d.append(fila_d)
        rows_n.append(fila_n)
    df_d, df_n = pd.DataFrame(rows_d, columns=cols_tabla), pd.DataFrame(rows_n, columns=cols_tabla)
    renderizar(df_d, df_n, negrita=True, colores=True)

else:
    st.subheader(f"Resultados {','.join(map(str, anos))}")
    cols_tabla = ["Concepto"] + tiendas
    rows_d, rows_n = [], []
    for c in CONCEPTOS:
        fila_d, fila_n = [c], [c]
        for t in tiendas:
            df_t = df[(df["Año"].isin(anos)) & (df["Mes"].isin(meses_sel)) & (df["Departamento"]==t)]
            res_t = calcular_res(df_t)
            val = res_t.get(c, 0.0)
            fila_n.append(val)
            fila_d.append(fmt_pct(val) if c == "R. B." else fmt_eur(val))
        rows_d.append(fila_d)
        rows_n.append(fila_n)
    df_d, df_n = pd.DataFrame(rows_d, columns=cols_tabla), pd.DataFrame(rows_n, columns=cols_tabla)
    renderizar(df_d, df_n, negrita=True)
