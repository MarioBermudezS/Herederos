import io
import pandas as pd
import streamlit as st
from typing import Dict, List

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

CSS_ESTILOS = """<style>#MainMenu{visibility:hidden;}footer{visibility:hidden;}.viewerBadge_container{display:none!important;}.block-container{padding:0.3rem 0.5rem!important;max-width:100%!important;}h1{font-size:0.95rem!important;margin:0!important;}h3{font-size:0.8rem!important;margin:0!important;}.stDataFrame{font-size:11px!important;}table{font-size:11px!important;line-height:1.2!important;}th,td{padding:2px 4px!important;}</style>"""

MESES_ORDEN = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
CONCEPTOS_KPI = ["Ventas", "Coste Ventas", "MARGEN BRUTO", "R. B.", "Otros Ingresos", "Ingresos Operativos", "Gastos Personal", "Alquileres", "Reparaciones", "Seguros", "Suministros", "Otros Servicios", "TOTAL GASTOS OPERATIVOS", "Amortizaciones", "GASTOS ESTRUCTURA", "B.A.I.I.", "Gastos Financieros", "Ingresos Financieros", "RDO. FINANCIERO", "Resultados Extraordinarios", "B.A.I."]
FILAS_NEGRITA_RESULTADOS = ["MARGEN BRUTO", "Ingresos Operativos", "TOTAL GASTOS OPERATIVOS", "GASTOS ESTRUCTURA", "B.A.I.I.", "RDO. FINANCIERO", "B.A.I."]
FILAS_NEGRITA_RB = ["TOTAL GRUPO"]
FILAS_NEGRITA_KPI = ["MARGEN BRUTO", "Ingresos Operativos", "TOTAL GASTOS OPERATIVOS", "GASTOS ESTRUCTURA", "B.A.I.I.", "B.A.I."]

GASTOS = {"Coste Ventas", "Gastos Personal", "Alquileres", "Reparaciones", "Seguros", "Suministros", "Otros Servicios", "TOTAL GASTOS OPERATIVOS", "Amortizaciones", "GASTOS ESTRUCTURA", "Gastos Financieros"}
INGRESOS = {"Ventas", "MARGEN BRUTO", "R. B.", "Otros Ingresos", "Ingresos Operativos", "Ingresos Financieros", "B.A.I.I.", "RDO. FINANCIERO", "B.A.I."}

st.markdown(CSS_ESTILOS, unsafe_allow_html=True)
st.title("Control de Resultados - Herederos")

def formato_porcentaje(valor: float, decimales: int = 2) -> str:
    return f"{valor * 100:,.{decimales}f}%".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_moneda(valor: float, decimales: int = 2) -> str:
    return f"{valor:,.{decimales}f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_variacion_pp(valor: float) -> str:
    return f"{valor * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", ".")

@st.cache_data
def load_data() -> pd.DataFrame:
    return pd.read_excel("BaseDatos2026.xlsx", sheet_name="BS")

def obtener_filtro_datos(df: pd.DataFrame, ano: int, meses: List[str], departamentos: List[str] = None) -> pd.DataFrame:
    mascara = (df["Año"] == ano) & (df["Mes"].isin(meses))
    if departamentos:
        mascara &= (df["Departamento"].isin(departamentos))
    return df[mascara]

def calcular_rb_puro(df_filtrado: pd.DataFrame) -> float:
    if df_filtrado.empty:
        return 0.0
    total_margen = 0.0
    total_ventas = 0.0
    for _, group in df_filtrado.groupby(["Año", "Mes", "Departamento"]):
        ventas = group[group["Resultados"] == "Ventas"]["Importe D"].sum()
        rb_rows = group[group["Resultados"].str.strip().str.upper().isin(["R. B.", "R.B.", "R.B"])]["Importe D"]
        rb_val = rb_rows.iloc[0] if not rb_rows.empty else 0.0
        total_ventas += ventas
        total_margen += ventas * rb_val
    if total_ventas == 0:
        resumen = df_filtrado.groupby("Resultados")["Importe D"].sum().to_dict()
        return resumen.get("R. B.", 0.0)
    return total_margen / total_ventas

def calcular_resultados(df_filtrado: pd.DataFrame) -> Dict[str, float]:
    if df_filtrado.empty:
        return {c: 0.0 for c in CONCEPTOS_KPI}
    resumen = df_filtrado.groupby("Resultados")["Importe D"].sum().to_dict()
    def get_v(cat):
        return resumen.get(cat, 0.0)
    ventas = get_v("Ventas")
    total_margen = 0.0
    total_ventas_calc = 0.0
    for _, group in df_filtrado.groupby(["Año", "Mes", "Departamento"]):
        v_row = group[group["Resultados"] == "Ventas"]["Importe D"].sum()
        rb_rows = group[group["Resultados"].str.strip().str.upper().isin(["R. B.", "R.B.", "R.B"])]["Importe D"]
        rb_val = rb_rows.iloc[0] if not rb_rows.empty else 0.0
        total_ventas_calc += v_row
        total_margen += v_row * rb_val
    margen_bruto = total_margen
    r_bruta = (margen_bruto / total_ventas_calc) if total_ventas_calc != 0 else 0.0
    if total_ventas_calc == 0 and ventas != 0:
        r_bruta = get_v("R. B.")
        margen_bruto = r_bruta * ventas
    coste_ventas = ventas - margen_bruto
    otros_ingresos = get_v("Otros Ingresos")
    ingresos_operativos = margen_bruto + otros_ingresos
    gastos_personal = get_v("Gastos Personal")
    alquileres = get_v("Alquileres")
    reparaciones = get_v("Reparaciones")
    seguros = get_v("Seguros")
    suministros = get_v("Suministros")
    otros_servicios = get_v("Otros Servicios")
    total_gastos_operativos = alquileres + reparaciones + seguros + suministros + otros_servicios
    amortizaciones = get_v("Amortizaciones")
    gastos_estructura = total_gastos_operativos + gastos_personal + amortizaciones
    baii = ingresos_operativos - gastos_personal - total_gastos_operativos - amortizaciones
    gastos_financieros = get_v("Gastos Financieros")
    ingresos_financieros = get_v("Ingresos Financieros")
    rdo_financiero = gastos_financieros + ingresos_financieros
    resultados_extraordinarios = get_v("Resultados Extraordinarios")
    bai = baii - rdo_financiero - resultados_extraordinarios
    return {"Ventas": ventas, "Coste Ventas": coste_ventas, "MARGEN BRUTO": margen_bruto, "R. B.": r_bruta, "Otros Ingresos": otros_ingresos, "Ingresos Operativos": ingresos_operativos, "Gastos Personal": gastos_personal, "Alquileres": alquileres, "Reparaciones": reparaciones, "Seguros": seguros, "Suministros": suministros, "Otros Servicios": otros_servicios, "TOTAL GASTOS OPERATIVOS": total_gastos_operativos, "Amortizaciones": amortizaciones, "GASTOS ESTRUCTURA": gastos_estructura, "B.A.I.I.": baii, "Gastos Financieros": gastos_financieros, "Ingresos Financieros": ingresos_financieros, "RDO. FINANCIERO": rdo_financiero, "Resultados Extraordinarios": resultados_extraordinarios, "B.A.I.": bai}

def render_dataframe_table(df_display: pd.DataFrame, df_numerico: pd.DataFrame = None, filas_negrita: List[str] = None, con_colores: bool = False) -> None:
    def estilo_fila(row):
        if filas_negrita:
            primera_col = df_display.columns[0]
            valor_primera = df_display.iloc[row.name][primera_col]
            if valor_primera in filas_negrita:
                return ['font-weight: bold; background-color: #f0f0f0;'] * len(row)
        return [''] * len(row)
    styled_df = df_display.style.apply(estilo_fila, axis=1)
    
    if con_colores and df_numerico is not None:
        for col in df_numerico.columns[1:]:
            for row_idx in range(len(df_display)):
                concepto = df_display.iloc[row_idx, 0]
                valores = pd.to_numeric(df_numerico[col].iloc[1:], errors='coerce')
                if valores.notna().any():
                    if concepto in GASTOS and row_idx == valores.idxmin():
                        styled_df = styled_df.applymap(lambda x: 'background-color: #90EE90', subset=pd.IndexSlice[row_idx, col])
                    elif concepto in INGRESOS and row_idx == valores.idxmax():
                        styled_df = styled_df.applymap(lambda x: 'background-color: #90EE90', subset=pd.IndexSlice[row_idx, col])
    
    st.dataframe(styled_df, use_container_width=True, hide_index=True)

def descargar_excel(df: pd.DataFrame, nombre_hoja: str, nombre_archivo: str, etiqueta: str) -> None:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    st.download_button(label=f"📥 {etiqueta}", data=output.getvalue(), file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

try:
    df = load_data()
except Exception as e:
    st.error(f"Error: {e}")
    st.stop()

st.sidebar.header("Parámetros")
modulo_principal = st.sidebar.radio("Módulo", ["Cuenta de Resultados", "Análisis R.B.", "KPI"])

anos_disponibles = [2024, 2025, 2026]
if "Año" in df.columns:
    anos_excel = sorted(df["Año"].dropna().unique())
    anos_disponibles = [a for a in anos_disponibles if a in anos_excel]
if not anos_disponibles:
    anos_disponibles = [2026]

ano = st.sidebar.selectbox("Año", anos_disponibles)

meses_excel = df["Mes"].dropna().unique().tolist() if "Mes" in df.columns else ["Enero"]
meses_disponibles = [m for m in MESES_ORDEN if m in meses_excel]
if not meses_disponibles:
    meses_disponibles = meses_excel

meses_sel = st.sidebar.multiselect("Mes(es)", meses_disponibles, default=meses_disponibles[:1])

df_ano = df[df["Año"] == ano] if "Año" in df.columns else df
departamentos_disponibles = sorted(df_ano["Departamento"].dropna().unique()) if "Departamento" in df_ano.columns else []

if modulo_principal == "Análisis R.B.":
    st.sidebar.markdown("---")
    tipo_analisis_rb = st.sidebar.radio("Vista", ["Mensual", "Acumulada", "Interanual"])
    tiendas_rb = st.sidebar.multiselect("Tiendas", departamentos_disponibles, default=departamentos_disponibles)
    if not tiendas_rb or not meses_sel:
        st.warning("Selecciona tiendas y meses")
        st.stop()
    
    if tipo_analisis_rb == "Mensual":
        st.subheader(f"R.B. Mensual {ano}")
        columnas_tabla = ["Tienda"] + meses_sel + (["Prom."] if len(meses_sel) > 1 else [])
        filas_display = []
        for tienda in tiendas_rb:
            fila_d = [tienda]
            for mes in meses_sel:
                df_f = obtener_filtro_datos(df, ano, [mes], [tienda])
                val_rb = calcular_rb_puro(df_f)
                fila_d.append(formato_porcentaje(val_rb))
            if len(meses_sel) > 1:
                df_f = obtener_filtro_datos(df, ano, meses_sel, [tienda])
                fila_d.append(formato_porcentaje(calcular_rb_puro(df_f)))
            filas_display.append(fila_d)
        if len(tiendas_rb) > 1:
            fila_d = ["TOTAL"]
            for mes in meses_sel:
                df_f = obtener_filtro_datos(df, ano, [mes], tiendas_rb)
                fila_d.append(formato_porcentaje(calcular_rb_puro(df_f)))
            if len(meses_sel) > 1:
                df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas_rb)
                fila_d.append(formato_porcentaje(calcular_rb_puro(df_f)))
            filas_display.append(fila_d)
        df_res = pd.DataFrame(filas_display, columns=columnas_tabla)
        render_dataframe_table(df_res, filas_negrita=FILAS_NEGRITA_RB)
    
    elif tipo_analisis_rb == "Acumulada":
        nombre_m = ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} meses"
        st.subheader(f"R.B. Acumulada {nombre_m} {ano}")
        filas = []
        for tienda in tiendas_rb:
            df_f = obtener_filtro_datos(df, ano, meses_sel, [tienda])
            filas.append([tienda, formato_porcentaje(calcular_rb_puro(df_f))])
        if len(tiendas_rb) > 1:
            df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas_rb)
            filas.append(["TOTAL", formato_porcentaje(calcular_rb_puro(df_f))])
        df_res = pd.DataFrame(filas, columns=["Tienda", "R.B."])
        render_dataframe_table(df_res, filas_negrita=FILAS_NEGRITA_RB)
    
    else:
        ano_ant = ano - 1
        nombre_m = ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} m"
        st.subheader(f"R.B. {ano} vs {ano_ant} ({nombre_m})")
        filas = []
        for tienda in tiendas_rb:
            df_a = obtener_filtro_datos(df, ano, meses_sel, [tienda])
            df_p = obtener_filtro_datos(df, ano_ant, meses_sel, [tienda])
            va = calcular_rb_puro(df_a)
            vp = calcular_rb_puro(df_p)
            filas.append([tienda, formato_porcentaje(va), formato_porcentaje(vp), formato_variacion_pp(va - vp)])
        if len(tiendas_rb) > 1:
            df_a = obtener_filtro_datos(df, ano, meses_sel, tiendas_rb)
            df_p = obtener_filtro_datos(df, ano_ant, meses_sel, tiendas_rb)
            va = calcular_rb_puro(df_a)
            vp = calcular_rb_puro(df_p)
            filas.append(["TOTAL", formato_porcentaje(va), formato_porcentaje(vp), formato_variacion_pp(va - vp)])
        df_res = pd.DataFrame(filas, columns=["Tienda", f"{ano}", f"{ano_ant}", "Var"])
        render_dataframe_table(df_res, filas_negrita=FILAS_NEGRITA_RB)

elif modulo_principal == "KPI":
    modo = st.sidebar.radio("Análisis", ["Mensual", "Multi-Tienda", "Interanual"])
    if modo == "Multi-Tienda":
        tiendas = st.sidebar.multiselect("Tiendas", departamentos_disponibles, default=departamentos_disponibles[:2] if len(departamentos_disponibles) >= 2 else departamentos_disponibles)
    else:
        tipo = st.sidebar.radio("Tipo", ["Una tienda", "Varias tiendas"])
        if tipo == "Una tienda":
            tienda_s = st.sidebar.selectbox("Tienda", departamentos_disponibles)
            tiendas = [tienda_s] if tienda_s else []
        else:
            tiendas = st.sidebar.multiselect("Tiendas", departamentos_disponibles, default=departamentos_disponibles)
    if not tiendas or not meses_sel:
        st.warning("Selecciona tiendas y meses")
        st.stop()
    
    nombre_m = ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} m"
    
    if modo == "Interanual":
        st.subheader(f"KPI {ano} vs {ano-1} ({nombre_m})")
        df_a = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        df_p = obtener_filtro_datos(df, ano-1, meses_sel, tiendas)
        ra = calcular_resultados(df_a)
        rp = calcular_resultados(df_p)
        va_v = ra.get("Ventas", 1.0)
        vp_v = rp.get("Ventas", 1.0)
        filas_d = []
        filas_n = []
        for c in CONCEPTOS_KPI:
            val_a = ra.get(c, 0.0)
            val_p = rp.get(c, 0.0)
            if c == "R. B.":
                kpi_a, kpi_p = val_a, val_p
            else:
                kpi_a = (val_a / va_v) if va_v != 0 else 0.0
                kpi_p = (val_p / vp_v) if vp_v != 0 else 0.0
            var = kpi_a - kpi_p
            filas_n.append([c, kpi_a, kpi_p, var])
            filas_d.append([c, formato_porcentaje(kpi_a), formato_porcentaje(kpi_p), formato_variacion_pp(var)])
        df_d = pd.DataFrame(filas_d, columns=["Concepto", f"{ano}", f"{ano-1}", "Var"])
        df_n = pd.DataFrame(filas_n, columns=["Concepto", f"{ano}", f"{ano-1}", "Var"])
        render_dataframe_table(df_d, df_numerico=df_n, filas_negrita=FILAS_NEGRITA_KPI, con_colores=True)
    
    elif modo == "Multi-Tienda":
        st.subheader(f"KPI Multi-Tienda ({nombre_m} {ano})")
        rf = {}
        for t in tiendas:
            df_f = obtener_filtro_datos(df, ano, meses_sel, [t])
            rf[t] = calcular_resultados(df_f)
        df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        rf["Total"] = calcular_resultados(df_f)
        cols_e = tiendas + ["Total"]
        filas_d = []
        filas_n = []
        for c in CONCEPTOS_KPI:
            fila_d = [c]
            fila_n = [c]
            for t in cols_e:
                v_v = rf[t].get("Ventas", 1.0)
                val = rf[t].get(c, 0.0)
                kpi = val if c == "R. B." else (val / v_v if v_v != 0 else 0.0)
                fila_n.append(kpi)
                fila_d.append(formato_porcentaje(kpi))
            filas_d.append(fila_d)
            filas_n.append(fila_n)
        df_d = pd.DataFrame(filas_d, columns=["Concepto"] + cols_e)
        df_n = pd.DataFrame(filas_n, columns=["Concepto"] + cols_e)
        render_dataframe_table(df_d, df_numerico=df_n, filas_negrita=FILAS_NEGRITA_KPI, con_colores=True)
    
    else:
        st.subheader(f"KPI ({nombre_m} {ano})")
        rf = {}
        if len(tiendas) > 1:
            for t in tiendas:
                df_f = obtener_filtro_datos(df, ano, meses_sel, [t])
                rf[t] = calcular_resultados(df_f)
            df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas)
            rf["Total"] = calcular_resultados(df_f)
            cols_e = tiendas + ["Total"]
        else:
            for m in meses_sel:
                df_f = obtener_filtro_datos(df, ano, [m], tiendas)
                rf[m] = calcular_resultados(df_f)
            if len(meses_sel) > 1:
                df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas)
                rf["Total"] = calcular_resultados(df_f)
                cols_e = meses_sel + ["Total"]
            else:
                cols_e = meses_sel
        filas_d = []
        filas_n = []
        for c in CONCEPTOS_KPI:
            fila_d = [c]
            fila_n = [c]
            for t in cols_e:
                v_v = rf[t].get("Ventas", 1.0)
                val = rf[t].get(c, 0.0)
                kpi = val if c == "R. B." else (val / v_v if v_v != 0 else 0.0)
                fila_n.append(kpi)
                fila_d.append(formato_porcentaje(kpi))
            filas_d.append(fila_d)
            filas_n.append(fila_n)
        df_d = pd.DataFrame(filas_d, columns=["Concepto"] + cols_e)
        df_n = pd.DataFrame(filas_n, columns=["Concepto"] + cols_e)
        render_dataframe_table(df_d, df_numerico=df_n, filas_negrita=FILAS_NEGRITA_KPI, con_colores=True)

else:
    modo = st.sidebar.radio("Análisis", ["Mensual", "Multi-Tienda", "Interanual"])
    if modo == "Multi-Tienda":
        tiendas = st.sidebar.multiselect("Tiendas", departamentos_disponibles, default=departamentos_disponibles[:2] if len(departamentos_disponibles) >= 2 else departamentos_disponibles)
    else:
        tipo = st.sidebar.radio("Tipo", ["Una tienda", "Varias tiendas"])
        if tipo == "Una tienda":
            tienda_s = st.sidebar.selectbox("Tienda", departamentos_disponibles)
            tiendas = [tienda_s] if tienda_s else []
        else:
            tiendas = st.sidebar.multiselect("Tiendas", departamentos_disponibles, default=departamentos_disponibles)
    if not tiendas or not meses_sel:
        st.warning("Selecciona tiendas y meses")
        st.stop()
    
    nombre_m = ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} m"
    
    if modo == "Interanual":
        st.subheader(f"Resultados {ano} vs {ano-1} ({nombre_m})")
        df_a = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        df_p = obtener_filtro_datos(df, ano-1, meses_sel, tiendas)
        ra = calcular_resultados(df_a)
        rp = calcular_resultados(df_p)
        filas_d = []
        filas_n = []
        for c in CONCEPTOS_KPI:
            val_a = ra.get(c, 0.0)
            val_p = rp.get(c, 0.0)
            if c == "R. B.":
                var_e = val_a - val_p
                filas_n.append([c, val_a, val_p, var_e, 0.0])
                filas_d.append([c, formato_porcentaje(val_a), formato_porcentaje(val_p), formato_variacion_pp(var_e), "-"])
            else:
                var_e = val_a - val_p
                var_pc = (var_e / abs(val_p) * 100) if val_p != 0 else 0.0
                filas_n.append([c, val_a, val_p, var_e, var_pc])
                filas_d.append([c, formato_moneda(val_a), formato_moneda(val_p), formato_moneda(var_e), formato_porcentaje(var_pc / 100) if var_pc != 0 else "0,00%"])
        df_d = pd.DataFrame(filas_d, columns=["Concepto", f"{ano}", f"{ano-1}", "Var€", "Var%"])
        df_n = pd.DataFrame(filas_n, columns=["Concepto", f"{ano}", f"{ano-1}", "Var€", "Var%"])
        render_dataframe_table(df_d, df_numerico=df_n, filas_negrita=FILAS_NEGRITA_RESULTADOS)
        descargar_excel(df_n, "Resultados", f"Resultados_{ano}.xlsx", "📥 Descargar")
    
    elif modo == "Multi-Tienda":
        st.subheader(f"Resultados ({nombre_m} {ano})")
        rf = {}
        for t in tiendas:
            df_f = obtener_filtro_datos(df, ano, meses_sel, [t])
            rf[t] = calcular_resultados(df_f)
        df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        rf["Total"] = calcular_resultados(df_f)
        cols_e = tiendas + ["Total"]
        filas_d = []
        filas_n = []
        for c in CONCEPTOS_KPI:
            fila_d = [c]
            fila_n = [c]
            for t in cols_e:
                val = rf[t].get(c, 0.0)
                fila_n.append(val)
                fila_d.append(formato_porcentaje(val) if c == "R. B." else formato_moneda(val))
            filas_d.append(fila_d)
            filas_n.append(fila_n)
        df_d = pd.DataFrame(filas_d, columns=["Concepto"] + cols_e)
        df_n = pd.DataFrame(filas_n, columns=["Concepto"] + cols_e)
        render_dataframe_table(df_d, df_numerico=df_n, filas_negrita=FILAS_NEGRITA_RESULTADOS)
        descargar_excel(df_n, "Resultados", f"Resultados_{ano}.xlsx", "📥 Descargar")
    
    else:
        st.subheader(f"Resultados ({nombre_m} {ano})")
        rf = {}
        if len(tiendas) > 1:
            for t in tiendas:
                df_f = obtener_filtro_datos(df, ano, meses_sel, [t])
                rf[t] = calcular_resultados(df_f)
            df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas)
            rf["Total"] = calcular_resultados(df_f)
            cols_e = tiendas + ["Total"]
        else:
            for m in meses_sel:
                df_f = obtener_filtro_datos(df, ano, [m], tiendas)
                rf[m] = calcular_resultados(df_f)
            if len(meses_sel) > 1:
                df_f = obtener_filtro_datos(df, ano, meses_sel, tiendas)
                rf["Total"] = calcular_resultados(df_f)
                cols_e = meses_sel + ["Total"]
            else:
                cols_e = meses_sel
        filas_d = []
        filas_n = []
        for c in CONCEPTOS_KPI:
            fila_d = [c]
            fila_n = [c]
            for t in cols_e:
                val = rf[t].get(c, 0.0)
                fila_n.append(val)
                fila_d.append(formato_porcentaje(val) if c == "R. B." else formato_moneda(val))
            filas_d.append(fila_d)
            filas_n.append(fila_n)
        df_d = pd.DataFrame(filas_d, columns=["Concepto"] + cols_e)
        df_n = pd.DataFrame(filas_n, columns=["Concepto"] + cols_e)
        render_dataframe_table(df_d, df_numerico=df_n, filas_negrita=FILAS_NEGRITA_RESULTADOS)
        descargar_excel(df_n, "Resultados", f"Resultados_{ano}.xlsx", "📥 Descargar")
