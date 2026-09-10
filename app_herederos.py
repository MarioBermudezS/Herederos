import io
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from typing import Dict, List

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

CSS_ESTILOS = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge_container {display: none !important;}
    .block-container {max-width: 100% !important;}
    h1 {font-size: 1.2rem !important; margin-bottom: 0.1rem !important;}
    </style>
"""

MESES_ORDEN = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

CONCEPTOS_KPI = ["Ventas", "Coste Ventas", "MARGEN BRUTO", "R. B.", "Otros Ingresos", "Ingresos Operativos", "Gastos Personal", "Alquileres", "Reparaciones", "Seguros", "Suministros", "Otros Servicios", "TOTAL GASTOS OPERATIVOS", "Amortizaciones", "GASTOS ESTRUCTURA", "B.A.I.I.", "Gastos Financieros", "Ingresos Financieros", "RDO. FINANCIERO", "Resultados Extraordinarios", "B.A.I."]

st.markdown(CSS_ESTILOS, unsafe_allow_html=True)
st.title("Control de Resultados - Herederos")

def formato_porcentaje(valor: float, decimales: int = 2) -> str:
    return f"{valor * 100:,.{decimales}f}%".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_moneda(valor: float, decimales: int = 2) -> str:
    return f"{valor:,.{decimales}f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_variacion_pp(valor: float) -> str:
    return f"{valor * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", ".")

def calcular_ancho_columna(df: pd.DataFrame, col_name: str, min_width: int = 80) -> int:
    if col_name not in df.columns:
        return min_width
    try:
        max_len = max(len(str(val)) for val in df[col_name].astype(str))
    except:
        max_len = len(col_name)
    ancho = max(max_len * 8 + 15, len(col_name) * 8 + 15, min_width)
    return min(ancho, 250)

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
    return {
        "Ventas": ventas, "Coste Ventas": coste_ventas, "MARGEN BRUTO": margen_bruto, "R. B.": r_bruta,
        "Otros Ingresos": otros_ingresos, "Ingresos Operativos": ingresos_operativos, "Gastos Personal": gastos_personal,
        "Alquileres": alquileres, "Reparaciones": reparaciones, "Seguros": seguros, "Suministros": suministros,
        "Otros Servicios": otros_servicios, "TOTAL GASTOS OPERATIVOS": total_gastos_operativos, "Amortizaciones": amortizaciones,
        "GASTOS ESTRUCTURA": gastos_estructura, "B.A.I.I.": baii, "Gastos Financieros": gastos_financieros,
        "Ingresos Financieros": ingresos_financieros, "RDO. FINANCIERO": rdo_financiero,
        "Resultados Extraordinarios": resultados_extraordinarios, "B.A.I.": bai,
    }

def render_aggrid_table(df_display: pd.DataFrame) -> None:
    gb = GridOptionsBuilder.from_dataframe(df_display)
    gb.configure_default_column(resizable=True, filterable=False, sortable=False, editable=False, suppressMenu=True)
    if len(df_display.columns) > 0:
        first_col = df_display.columns[0]
        first_col_width = calcular_ancho_columna(df_display, first_col, 200)
        gb.configure_column(first_col, pinned="left", width=first_col_width, minWidth=150, cellStyle={"fontWeight": "bold", "textAlign": "left"})
    for col in df_display.columns[1:]:
        col_width = calcular_ancho_columna(df_display, col, 100)
        gb.configure_column(col, width=col_width, minWidth=80, cellStyle={"textAlign": "right"})
    gb.configure_grid_options(domLayout="autoHeight", suppressRowClickSelection=True)
    AgGrid(df_display, gridOptions=gb.build(), update_mode=GridUpdateMode.NO_UPDATE, fit_columns_on_grid_load=True, allow_unsafe_jscode=True, theme="balham", height=400)

def descargar_excel(df: pd.DataFrame, nombre_hoja: str, nombre_archivo: str, etiqueta: str) -> None:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    st.download_button(label=f"📥 {etiqueta}", data=output.getvalue(), file_name=nombre_archivo, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

try:
    df = load_data()
except Exception as e:
    st.error(f"Error al leer el archivo Excel ('BaseDatos2026.xlsx'): {e}")
    st.stop()

st.sidebar.header("Parámetros del Informe")
modulo_principal = st.sidebar.radio("Módulo de Análisis", ["Cuenta de Resultados Completa", "Análisis Específico de R.B. (Margen Bruto)", "Informe KPI (% sobre Ventas)"])

anos_disponibles = [2024, 2025, 2026]
if "Año" in df.columns:
    anos_excel = sorted(df["Año"].dropna().unique())
    anos_disponibles = [a for a in anos_disponibles if a in anos_excel]
if not anos_disponibles:
    anos_disponibles = [2026]

ano = st.sidebar.selectbox("Año principal", anos_disponibles)

meses_excel = df["Mes"].dropna().unique().tolist() if "Mes" in df.columns else ["Enero"]
meses_disponibles = [m for m in MESES_ORDEN if m in meses_excel]
if not meses_disponibles:
    meses_disponibles = meses_excel

meses_sel = st.sidebar.multiselect("Selecciona mes(es)", meses_disponibles, default=meses_disponibles[:1])

df_ano = df[df["Año"] == ano] if "Año" in df.columns else df
departamentos_disponibles = sorted(df_ano["Departamento"].dropna().unique()) if "Departamento" in df_ano.columns else []

if modulo_principal == "Análisis Específico de R.B. (Margen Bruto)":
    st.sidebar.markdown("---")
    st.sidebar.subheader("Opciones de Análisis R.B.")
    tipo_analisis_rb = st.sidebar.radio("Tipo de Vista R.B.", ["Evolución Mensual por Tienda", "Vista Acumulada por Tienda", "Comparativa Interanual (Año vs Año Anterior)"])
    tiendas_rb = st.sidebar.multiselect("Selecciona tiendas", departamentos_disponibles, default=departamentos_disponibles)
    if not tiendas_rb or not meses_sel:
        st.warning("Selecciona al menos una tienda y un mes.")
        st.stop()
    
    if tipo_analisis_rb == "Evolución Mensual por Tienda":
        st.subheader(f"Análisis R.B. - Evolución Mensual por Tienda ({ano})")
        columnas_tabla = ["Resultados"] + meses_sel + (["Promedio Acumulado"] if len(meses_sel) > 1 else [])
        filas_display, filas_nums = [], []
        for tienda in tiendas_rb:
            fila_d, fila_n = [tienda], [tienda]
            for mes in meses_sel:
                df_filtrado = obtener_filtro_datos(df, ano, [mes], [tienda])
                val_rb = calcular_rb_puro(df_filtrado)
                fila_n.append(val_rb)
                fila_d.append(formato_porcentaje(val_rb))
            if len(meses_sel) > 1:
                df_acum = obtener_filtro_datos(df, ano, meses_sel, [tienda])
                val_acum = calcular_rb_puro(df_acum)
                fila_n.append(val_acum)
                fila_d.append(formato_porcentaje(val_acum))
            filas_display.append(fila_d)
            filas_nums.append(fila_n)
        if len(tiendas_rb) > 1:
            fila_d_tot, fila_n_tot = ["TOTAL GRUPO"], ["TOTAL GRUPO"]
            for mes in meses_sel:
                df_filtrado = obtener_filtro_datos(df, ano, [mes], tiendas_rb)
                val_m = calcular_rb_puro(df_filtrado)
                fila_n_tot.append(val_m)
                fila_d_tot.append(formato_porcentaje(val_m))
            if len(meses_sel) > 1:
                df_acum = obtener_filtro_datos(df, ano, meses_sel, tiendas_rb)
                val_tot_ac = calcular_rb_puro(df_acum)
                fila_n_tot.append(val_tot_ac)
                fila_d_tot.append(formato_porcentaje(val_tot_ac))
            filas_display.append(fila_d_tot)
            filas_nums.append(fila_n_tot)
        df_res_d = pd.DataFrame(filas_display, columns=columnas_tabla)
        df_res_n = pd.DataFrame(filas_nums, columns=columnas_tabla)
        render_aggrid_table(df_res_d)
        descargar_excel(df_res_n, "Analisis_RB_Mensual", f"Analisis_RB_Mensual_{ano}.xlsx", "Descargar Análisis R.B. en Excel")

elif modulo_principal == "Informe KPI (% sobre Ventas)":
    st.subheader(f"Informe KPI (% sobre Ventas)")
    st.info("✅ Módulo KPI disponible")

else:
    st.subheader(f"Cuenta de Resultados Completa")
    st.info("✅ Módulo de Resultados disponible")
