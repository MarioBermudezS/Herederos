import io
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode
from typing import Dict, List, Tuple

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

# =====================================================================
# CONFIGURACIÓN Y ESTILOS
# =====================================================================

CSS_ESTILOS = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge_container {display: none !important;}
    a[href*="github.com"] {display: none !important;}
    
    .block-container {
        padding-top: 0.5rem !important;
        padding-bottom: 0.5rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    h1 {font-size: 1.20rem !important; margin: 0 0 0.10rem 0 !important;}
    h2 {font-size: 1.02rem !important; margin: 0.10rem 0 !important;}
    h3 {font-size: 0.95rem !important; margin: 0.10rem 0 !important;}

    /* Interfaz más compacta para aprovechar toda la pantalla */
    div[data-testid="stVerticalBlock"] {gap: 0.25rem !important;}
    div[data-testid="stSidebar"] .block-container {
        padding-top: 0.45rem !important;
        padding-bottom: 0.35rem !important;
    }
    div[data-testid="stSidebar"] .stRadio,
    div[data-testid="stSidebar"] .stSelectbox,
    div[data-testid="stSidebar"] .stMultiSelect {
        margin-bottom: 0.10rem !important;
    }
    div[data-testid="stDownloadButton"] {margin-top: 0.15rem !important;}
    </style>
"""

MESES_ORDEN = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

CONCEPTOS_KPI = [
    "Ventas",
    "Coste Ventas",
    "MARGEN BRUTO",
    "R. B.",
    "Otros Ingresos",
    "Ingresos Operativos",
    "Gastos Personal",
    "Alquileres",
    "Reparaciones",
    "Seguros",
    "Suministros",
    "Otros Servicios",
    "TOTAL GASTOS OPERATIVOS",
    "Amortizaciones",
    "GASTOS ESTRUCTURA",
    "B.A.I.I.",
    "Gastos Financieros",
    "Ingresos Financieros",
    "RDO. FINANCIERO",
    "Resultados Extraordinarios",
    "B.A.I.",
]

st.markdown(CSS_ESTILOS, unsafe_allow_html=True)
st.title("Control de Resultados - Herederos")

# =====================================================================
# FUNCIONES AUXILIARES DE FORMATO Y UTILIDAD
# =====================================================================

def formato_porcentaje(valor: float, decimales: int = 2) -> str:
    """Convierte un decimal a string formateado como porcentaje con locale español."""
    return f"{valor * 100:,.{decimales}f}%".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_moneda(valor: float, decimales: int = 2) -> str:
    """Convierte un número a string formateado como euros con locale español."""
    return f"{valor:,.{decimales}f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_variacion_pp(valor: float) -> str:
    """Formatea variaciones en puntos porcentuales."""
    return f"{valor * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", ".")

def calcular_ancho_columna(df: pd.DataFrame, col_name: str, min_width: int = 80) -> int:
    """
    Calcula el ancho óptimo de una columna basado en su contenido.
    
    SOLUCIÓN 1: Ancho dinámico por contenido
    """
    # Ancho mínimo garantizado
    if col_name not in df.columns:
        return min_width
    
    # Obtener el valor más largo como string
    try:
        max_len = max(
            len(str(val)) for val in df[col_name].astype(str)
        )
    except:
        max_len = len(col_name)
    
    # Calcular ancho: 8 píxeles por carácter + padding
    ancho = max(max_len * 8 + 15, len(col_name) * 8 + 15, min_width)
    
    # Cap máximo para evitar columnas gigantes
    return min(ancho, 250)

@st.cache_data
def load_data() -> pd.DataFrame:
    """Carga datos del archivo Excel."""
    return pd.read_excel("BaseDatos2026.xlsx", sheet_name="BS")

def obtener_filtro_datos(
    df: pd.DataFrame, 
    ano: int, 
    meses: List[str], 
    departamentos: List[str] = None
) -> pd.DataFrame:
    """Filtra el dataframe por año, meses y departamentos."""
    mascara = (df["Año"] == ano) & (df["Mes"].isin(meses))
    if departamentos:
        mascara &= (df["Departamento"].isin(departamentos))
    return df[mascara]

# =====================================================================
# FUNCIONES DE CÁLCULO
# =====================================================================

def calcular_rb_puro(df_filtrado: pd.DataFrame) -> float:
    """Calcula el margen bruto puro (R.B.) como porcentaje."""
    if df_filtrado.empty:
        return 0.0
    
    total_margen = 0.0
    total_ventas = 0.0
    
    for _, group in df_filtrado.groupby(["Año", "Mes", "Departamento"]):
        ventas = group[group["Resultados"] == "Ventas"]["Importe D"].sum()
        rb_rows = group[
            group["Resultados"].str.strip().str.upper().isin(["R. B.", "R.B.", "R.B"])
        ]["Importe D"]
        rb_val = rb_rows.iloc[0] if not rb_rows.empty else 0.0
        
        total_ventas += ventas
        total_margen += ventas * rb_val
    
    if total_ventas == 0:
        resumen = df_filtrado.groupby("Resultados")["Importe D"].sum().to_dict()
        return resumen.get("R. B.", 0.0)
    
    return total_margen / total_ventas

def calcular_resultados(df_filtrado: pd.DataFrame) -> Dict[str, float]:
    """Calcula todos los conceptos de la cuenta de resultados."""
    if df_filtrado.empty:
        return {c: 0.0 for c in CONCEPTOS_KPI}
    
    resumen = df_filtrado.groupby("Resultados")["Importe D"].sum().to_dict()
    
    def get_v(cat):
        return resumen.get(cat, 0.0)
    
    # Cálculos principales
    ventas = get_v("Ventas")
    
    # Margen bruto
    total_margen = 0.0
    total_ventas_calc = 0.0
    for _, group in df_filtrado.groupby(["Año", "Mes", "Departamento"]):
        v_row = group[group["Resultados"] == "Ventas"]["Importe D"].sum()
        rb_rows = group[
            group["Resultados"].str.strip().str.upper().isin(["R. B.", "R.B.", "R.B"])
        ]["Importe D"]
        rb_val = rb_rows.iloc[0] if not rb_rows.empty else 0.0
        total_ventas_calc += v_row
        total_margen += v_row * rb_val
    
    margen_bruto = total_margen
    r_bruta = (margen_bruto / total_ventas_calc) if total_ventas_calc != 0 else 0.0
    
    if total_ventas_calc == 0 and ventas != 0:
        r_bruta = get_v("R. B.")
        margen_bruto = r_bruta * ventas
    
    coste_ventas = ventas - margen_bruto
    
    # Estructura de costos
    otros_ingresos = get_v("Otros Ingresos")
    ingresos_operativos = margen_bruto + otros_ingresos
    
    gastos_personal = get_v("Gastos Personal")
    alquileres = get_v("Alquileres")
    reparaciones = get_v("Reparaciones")
    seguros = get_v("Seguros")
    suministros = get_v("Suministros")
    otros_servicios = get_v("Otros Servicios")
    
    total_gastos_operativos = (
        alquileres + reparaciones + seguros + suministros + otros_servicios
    )
    amortizaciones = get_v("Amortizaciones")
    gastos_estructura = total_gastos_operativos + gastos_personal + amortizaciones
    
    # BAII
    baii = (
        ingresos_operativos
        - gastos_personal
        - total_gastos_operativos
        - amortizaciones
    )
    
    # Resultado financiero
    gastos_financieros = get_v("Gastos Financieros")
    ingresos_financieros = get_v("Ingresos Financieros")
    rdo_financiero = gastos_financieros + ingresos_financieros
    
    resultados_extraordinarios = get_v("Resultados Extraordinarios")
    bai = baii - rdo_financiero - resultados_extraordinarios
    
    return {
        "Ventas": ventas,
        "Coste Ventas": coste_ventas,
        "MARGEN BRUTO": margen_bruto,
        "R. B.": r_bruta,
        "Otros Ingresos": otros_ingresos,
        "Ingresos Operativos": ingresos_operativos,
        "Gastos Personal": gastos_personal,
        "Alquileres": alquileres,
        "Reparaciones": reparaciones,
        "Seguros": seguros,
        "Suministros": suministros,
        "Otros Servicios": otros_servicios,
        "TOTAL GASTOS OPERATIVOS": total_gastos_operativos,
        "Amortizaciones": amortizaciones,
        "GASTOS ESTRUCTURA": gastos_estructura,
        "B.A.I.I.": baii,
        "Gastos Financieros": gastos_financieros,
        "Ingresos Financieros": ingresos_financieros,
        "RDO. FINANCIERO": rdo_financiero,
        "Resultados Extraordinarios": resultados_extraordinarios,
        "B.A.I.": bai,
    }

# =====================================================================
# FUNCIONES DE RENDERIZACIÓN
# =====================================================================

def render_aggrid_table(df_display: pd.DataFrame, modo: str = "auto") -> None:
    """
    Renderiza la tabla completa en formato compacto.

    Objetivos:
    - Mostrar todas las filas sin scroll vertical interno.
    - Autoajustar cada columna según su cabecera y contenido.
    - Resaltar con negrita y sombreado las líneas principales.
    """
    if df_display.empty:
        st.info("No hay datos para mostrar con los filtros seleccionados.")
        return

    # Líneas contables que deben destacarse en toda la fila.
    filas_negrita = {
        "MARGEN BRUTO",
        "R. B.",
        "Ingresos Operativos",
        "TOTAL GASTOS OPERATIVOS",
        "GASTOS ESTRUCTURA",
        "B.A.I.I.",
        "RDO. FINANCIERO",
        "B.A.I.",
        "TOTAL GRUPO",
        "Total",
    }

    filas_js = ",".join(repr(x) for x in sorted(filas_negrita))
    get_row_style = JsCode(
        f"""
        function(params) {{
            const valor = params.data && params.data.Resultados ? String(params.data.Resultados) : '';
            const filasNegrita = [{filas_js}];
            if (filasNegrita.includes(valor)) {{
                return {{
                    'fontWeight': '700',
                    'backgroundColor': '#e9ecef'
                }};
            }}
            return null;
        }}
        """
    )

    gb = GridOptionsBuilder.from_dataframe(df_display)
    gb.configure_default_column(
        resizable=True,
        filterable=False,
        sortable=False,
        editable=False,
        suppressMenu=True,
        wrapText=False,
        autoHeight=False,
    )

    # AUTOAJUSTE REAL: cada columna se dimensiona por el texto más largo
    # entre la cabecera y los valores mostrados.
    for i, col in enumerate(df_display.columns):
        if i == 0:
            ancho = calcular_ancho_columna(df_display, col, 150)
            gb.configure_column(
                col,
                pinned="left",
                width=ancho,
                minWidth=150,
                maxWidth=280,
                cellStyle={"textAlign": "left"},
            )
        else:
            ancho = calcular_ancho_columna(df_display, col, 80)
            gb.configure_column(
                col,
                width=ancho,
                minWidth=80,
                maxWidth=250,
                cellStyle={"textAlign": "right"},
            )

    gb.configure_grid_options(
        domLayout="normal",
        suppressRowClickSelection=True,
        rowHeight=24,
        headerHeight=28,
        getRowStyle=get_row_style,
        suppressHorizontalScroll=False,
    )

    # Altura calculada para mostrar todas las filas sin scroll vertical interno.
    altura_tabla = 34 + (len(df_display) * 24)

    AgGrid(
        df_display,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        theme="balham",
        height=altura_tabla,
    )

def descargar_excel(df: pd.DataFrame, nombre_hoja: str, nombre_archivo: str, etiqueta: str) -> None:
    """Genera y descarga archivo Excel."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    
    st.download_button(
        label=f"📥 {etiqueta}",
        data=output.getvalue(),
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# =====================================================================
# CARGA DE DATOS E INICIALIZACIÓN
# =====================================================================

try:
    df = load_data()
except Exception as e:
    st.error(f"Error al leer el archivo Excel ('BaseDatos2026.xlsx'): {e}")
    st.stop()

# Sidebar - Filtros principales
st.sidebar.header("Parámetros del Informe")

modulo_principal = st.sidebar.radio(
    "Módulo de Análisis",
    [
        "Cuenta de Resultados Completa",
        "Análisis Específico de R.B. (Margen Bruto)",
        "Informe KPI (% sobre Ventas)",
    ],
)

# Años disponibles
anos_disponibles = [2024, 2025, 2026]
if "Año" in df.columns:
    anos_excel = sorted(df["Año"].dropna().unique())
    anos_disponibles = [a for a in anos_disponibles if a in anos_excel]
if not anos_disponibles:
    anos_disponibles = [2026]

ano = st.sidebar.selectbox("Año principal", anos_disponibles)

# Meses disponibles
meses_excel = df["Mes"].dropna().unique().tolist() if "Mes" in df.columns else ["Enero"]
meses_disponibles = [m for m in MESES_ORDEN if m in meses_excel]
if not meses_disponibles:
    meses_disponibles = meses_excel

meses_sel = st.sidebar.multiselect(
    "Selecciona mes(es)", meses_disponibles, default=meses_disponibles[:1]
)

# Departamentos disponibles
df_ano = df[df["Año"] == ano] if "Año" in df.columns else df
departamentos_disponibles = (
    sorted(df_ano["Departamento"].dropna().unique())
    if "Departamento" in df_ano.columns
    else []
)

# =====================================================================
# MÓDULO 1: ANÁLISIS ESPECÍFICO DE R.B. (MARGEN BRUTO)
# =====================================================================

if modulo_principal == "Análisis Específico de R.B. (Margen Bruto)":
    st.sidebar.markdown("---")
    st.sidebar.subheader("Opciones de Análisis R.B.")
    
    tipo_analisis_rb = st.sidebar.radio(
        "Tipo de Vista R.B.",
        [
            "Evolución Mensual por Tienda",
            "Vista Acumulada por Tienda",
            "Comparativa Interanual (Año vs Año Anterior)",
        ],
    )
    
    tiendas_rb = st.sidebar.multiselect(
        "Selecciona tiendas",
        departamentos_disponibles,
        default=departamentos_disponibles,
    )
    
    if not tiendas_rb or not meses_sel:
        st.warning("Selecciona al menos una tienda y un mes.")
        st.stop()
    
    # =====================================================================
    # EVOLUCIÓN MENSUAL POR TIENDA
    # =====================================================================
    if tipo_analisis_rb == "Evolución Mensual por Tienda":
        st.subheader(f"Análisis R.B. - Evolución Mensual por Tienda ({ano})")
        
        columnas_tabla = ["Resultados"] + meses_sel
        if len(meses_sel) > 1:
            columnas_tabla.append("Promedio Acumulado")
        
        filas_display = []
        filas_nums = []
        
        for tienda in tiendas_rb:
            fila_d = [tienda]
            fila_n = [tienda]
            
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
        
        # TOTAL GRUPO
        if len(tiendas_rb) > 1:
            fila_d_tot = ["TOTAL GRUPO"]
            fila_n_tot = ["TOTAL GRUPO"]
            
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
        
        render_aggrid_table(df_res_d, modo="auto")
        descargar_excel(df_res_n, "Analisis_RB_Mensual", f"Analisis_RB_Mensual_{ano}.xlsx", 
                       "Descargar Análisis R.B. en Excel")
    
    # =====================================================================
    # VISTA ACUMULADA POR TIENDA
    # =====================================================================
    elif tipo_analisis_rb == "Vista Acumulada por Tienda":
        nombre_m_str = (
            ", ".join(meses_sel)
            if len(meses_sel) <= 3
            else f"{len(meses_sel)} meses acumulados"
        )
        st.subheader(f"Análisis R.B. - Vista Acumulada ({nombre_m_str} {ano})")
        
        filas_acum_d = []
        filas_acum_n = []
        
        for tienda in tiendas_rb:
            df_filtrado = obtener_filtro_datos(df, ano, meses_sel, [tienda])
            val_rb = calcular_rb_puro(df_filtrado)
            filas_acum_d.append([tienda, formato_porcentaje(val_rb)])
            filas_acum_n.append([tienda, val_rb])
        
        if len(tiendas_rb) > 1:
            df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas_rb)
            val_tot = calcular_rb_puro(df_filtrado)
            filas_acum_d.append(["TOTAL GRUPO", formato_porcentaje(val_tot)])
            filas_acum_n.append(["TOTAL GRUPO", val_tot])
        
        df_acum_d = pd.DataFrame(filas_acum_d, columns=["Resultados", f"Acumulado {nombre_m_str}"])
        df_acum_n = pd.DataFrame(filas_acum_n, columns=["Resultados", f"Acumulado {nombre_m_str}"])
        
        render_aggrid_table(df_acum_d, modo="auto")
        descargar_excel(df_acum_n, "Analisis_RB_Acumulado", f"Analisis_RB_Acumulado_{ano}.xlsx",
                       "Descargar Acumulado R.B. en Excel")
    
    # =====================================================================
    # COMPARATIVA INTERANUAL
    # =====================================================================
    else:
        ano_ant = ano - 1
        nombre_m_str = (
            ", ".join(meses_sel)
            if len(meses_sel) <= 3
            else f"{len(meses_sel)} meses"
        )
        st.subheader(f"Comparativa Interanual R.B. ({nombre_m_str}): {ano} vs {ano_ant}")
        
        columnas_interanual_rb = ["Resultados", f"R.B. {ano}", f"R.B. {ano_ant}", "Var. pp"]
        filas_inter_d = []
        filas_inter_n = []
        
        for tienda in tiendas_rb:
            df_act = obtener_filtro_datos(df, ano, meses_sel, [tienda])
            df_ant = obtener_filtro_datos(df, ano_ant, meses_sel, [tienda])
            
            val_act = calcular_rb_puro(df_act)
            val_ant = calcular_rb_puro(df_ant)
            var_pp = val_act - val_ant
            
            filas_inter_n.append([tienda, val_act, val_ant, var_pp])
            filas_inter_d.append([
                tienda,
                formato_porcentaje(val_act),
                formato_porcentaje(val_ant),
                formato_variacion_pp(var_pp),
            ])
        
        if len(tiendas_rb) > 1:
            df_act = obtener_filtro_datos(df, ano, meses_sel, tiendas_rb)
            df_ant = obtener_filtro_datos(df, ano_ant, meses_sel, tiendas_rb)
            
            val_tot_act = calcular_rb_puro(df_act)
            val_tot_ant = calcular_rb_puro(df_ant)
            var_tot_pp = val_tot_act - val_tot_ant
            
            filas_inter_n.append(["TOTAL GRUPO", val_tot_act, val_tot_ant, var_tot_pp])
            filas_inter_d.append([
                "TOTAL GRUPO",
                formato_porcentaje(val_tot_act),
                formato_porcentaje(val_tot_ant),
                formato_variacion_pp(var_tot_pp),
            ])
        
        df_inter_d = pd.DataFrame(filas_inter_d, columns=columnas_interanual_rb)
        df_inter_n = pd.DataFrame(filas_inter_n, columns=columnas_interanual_rb)
        
        render_aggrid_table(df_inter_d, modo="auto")
        descargar_excel(df_inter_n, "Interanual_RB", f"Comparativa_Interanual_RB_{ano}.xlsx",
                       "Descargar Comparativa R.B. en Excel")

# =====================================================================
# MÓDULO 2: INFORME KPI (% SOBRE VENTAS)
# =====================================================================

elif modulo_principal == "Informe KPI (% sobre Ventas)":
    modo_analisis = st.sidebar.radio(
        "Tipo de Análisis KPI",
        [
            "Evolución Mensual / Tienda",
            "Comparativa Multi-Tienda (Totales)",
            "Comparativa Interanual (Año vs Año Anterior)",
        ],
    )
    
    # Selección de tiendas/meses
    if modo_analisis == "Comparativa Multi-Tienda (Totales)":
        tiendas = st.sidebar.multiselect(
            "Selecciona tiendas a comparar",
            departamentos_disponibles,
            default=departamentos_disponibles[:2]
            if len(departamentos_disponibles) >= 2
            else departamentos_disponibles,
        )
    else:
        tipo_consulta = st.sidebar.radio(
            "Tipo de consulta", ["Una tienda", "Conjunto de tiendas"]
        )
        if tipo_consulta == "Una tienda":
            tienda_sel = st.sidebar.selectbox("Selecciona tienda", departamentos_disponibles)
            tiendas = [tienda_sel] if tienda_sel else []
        else:
            tiendas = st.sidebar.multiselect(
                "Selecciona tiendas",
                departamentos_disponibles,
                default=departamentos_disponibles,
            )
    
    if not tiendas or not meses_sel:
        st.warning("Selecciona al menos una tienda y un mes.")
        st.stop()
    
    nombre_meses_str = (
        ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} meses"
    )
    
    if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":
        st.subheader(
            f"Informe KPI (% sobre Ventas) - Interanual: {nombre_meses_str} ({ano} vs {ano - 1})"
        )
        
        ano_anterior = ano - 1
        datos_fuente = {}
        
        df_ant = obtener_filtro_datos(df, ano_anterior, meses_sel, tiendas)
        df_act = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        
        datos_fuente["Ant"] = calcular_resultados(df_ant)
        datos_fuente["Act"] = calcular_resultados(df_act)
        
        columnas_eje = [f"Total {ano}", f"Total {ano_anterior}", "Var. pp"]
        columnas_tabla = ["Resultados"] + columnas_eje
        filas_tabla_display = []
        filas_valores_numericos = []
        
        for concepto in CONCEPTOS_KPI:
            fila_disp = [concepto]
            fila_num = [concepto]
            
            res_ant = datos_fuente["Ant"]
            res_act = datos_fuente["Act"]
            v_ventas_act = res_act.get("Ventas", 1.0)
            v_ventas_ant = res_ant.get("Ventas", 1.0)
            
            val_act_abs = res_act.get(concepto, 0.0)
            val_ant_abs = res_ant.get(concepto, 0.0)
            
            if concepto == "R. B.":
                kpi_act = val_act_abs
                kpi_ant = val_ant_abs
            else:
                kpi_act = (val_act_abs / v_ventas_act) if v_ventas_act != 0 else 0.0
                kpi_ant = (val_ant_abs / v_ventas_ant) if v_ventas_ant != 0 else 0.0
            
            var_pp = kpi_act - kpi_ant
            fila_num.extend([kpi_act, kpi_ant, var_pp])
            fila_disp.append(formato_porcentaje(kpi_act))
            fila_disp.append(formato_porcentaje(kpi_ant))
            fila_disp.append(formato_variacion_pp(var_pp))
            
            filas_tabla_display.append(fila_disp)
            filas_valores_numericos.append(fila_num)
        
        df_kpi_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
        df_kpi_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)
        
        render_aggrid_table(df_kpi_display, modo="auto")
        descargar_excel(df_kpi_numericos, "Informe_KPI", f"Informe_KPI_Ventas_{ano}.xlsx",
                       "Descargar Informe KPI en Excel")
    
    elif modo_analisis == "Comparativa Multi-Tienda (Totales)":
        st.subheader(f"Informe KPI (% sobre Ventas) - Multi-Tienda ({nombre_meses_str} {ano})")
        
        datos_fuente = {}
        for tienda in tiendas:
            df_filtrado = obtener_filtro_datos(df, ano, meses_sel, [tienda])
            datos_fuente[tienda] = calcular_resultados(df_filtrado)
        
        # Totales
        df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        datos_fuente["Total"] = calcular_resultados(df_filtrado)
        
        columnas_eje = tiendas + ["Total"]
        columnas_tabla = ["Resultados"] + columnas_eje
        filas_tabla_display = []
        filas_valores_numericos = []
        
        for concepto in CONCEPTOS_KPI:
            fila_disp = [concepto]
            fila_num = [concepto]
            
            for item in columnas_eje:
                res_item = datos_fuente[item]
                v_ventas_item = res_item.get("Ventas", 1.0)
                val_abs = res_item.get(concepto, 0.0)
                
                if concepto == "R. B.":
                    kpi_val = val_abs
                else:
                    kpi_val = (val_abs / v_ventas_item) if v_ventas_item != 0 else 0.0
                
                fila_num.append(kpi_val)
                fila_disp.append(formato_porcentaje(kpi_val))
            
            filas_tabla_display.append(fila_disp)
            filas_valores_numericos.append(fila_num)
        
        df_kpi_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
        df_kpi_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)
        
        render_aggrid_table(df_kpi_display, modo="auto")
        descargar_excel(df_kpi_numericos, "Informe_KPI", f"Informe_KPI_Ventas_{ano}.xlsx",
                       "Descargar Informe KPI en Excel")
    
    else:  # Evolución Mensual / Tienda
        st.subheader(f"Informe KPI (% sobre Ventas) - ({nombre_meses_str} {ano})")
        
        datos_fuente = {}
        
        if len(tiendas) > 1:
            # Comparativa por tiendas
            for tienda in tiendas:
                df_filtrado = obtener_filtro_datos(df, ano, meses_sel, [tienda])
                datos_fuente[tienda] = calcular_resultados(df_filtrado)
            
            df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas)
            datos_fuente["Total"] = calcular_resultados(df_filtrado)
            columnas_eje = tiendas + ["Total"]
        else:
            # Evolución por meses
            for mes in meses_sel:
                df_filtrado = obtener_filtro_datos(df, ano, [mes], tiendas)
                datos_fuente[mes] = calcular_resultados(df_filtrado)
            
            if len(meses_sel) > 1:
                df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas)
                datos_fuente["Total"] = calcular_resultados(df_filtrado)
                columnas_eje = meses_sel + ["Total"]
            else:
                columnas_eje = meses_sel
        
        columnas_tabla = ["Resultados"] + columnas_eje
        filas_tabla_display = []
        filas_valores_numericos = []
        
        for concepto in CONCEPTOS_KPI:
            fila_disp = [concepto]
            fila_num = [concepto]
            
            for item in columnas_eje:
                res_item = datos_fuente[item]
                v_ventas_item = res_item.get("Ventas", 1.0)
                val_abs = res_item.get(concepto, 0.0)
                
                if concepto == "R. B.":
                    kpi_val = val_abs
                else:
                    kpi_val = (val_abs / v_ventas_item) if v_ventas_item != 0 else 0.0
                
                fila_num.append(kpi_val)
                fila_disp.append(formato_porcentaje(kpi_val))
            
            filas_tabla_display.append(fila_disp)
            filas_valores_numericos.append(fila_num)
        
        df_kpi_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
        df_kpi_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)
        
        render_aggrid_table(df_kpi_display, modo="auto")
        descargar_excel(df_kpi_numericos, "Informe_KPI", f"Informe_KPI_Ventas_{ano}.xlsx",
                       "Descargar Informe KPI en Excel")

# =====================================================================
# MÓDULO 3: CUENTA DE RESULTADOS COMPLETA
# =====================================================================

else:
    modo_analisis = st.sidebar.radio(
        "Tipo de Análisis",
        [
            "Evolución Mensual / Tienda",
            "Comparativa Multi-Tienda (Totales)",
            "Comparativa Interanual (Año vs Año Anterior)",
        ],
    )
    
    # Selección de tiendas
    if modo_analisis == "Comparativa Multi-Tienda (Totales)":
        tiendas = st.sidebar.multiselect(
            "Selecciona tiendas a comparar",
            departamentos_disponibles,
            default=departamentos_disponibles[:2]
            if len(departamentos_disponibles) >= 2
            else departamentos_disponibles,
        )
    else:
        tipo_consulta = st.sidebar.radio(
            "Tipo de consulta", ["Una tienda", "Conjunto de tiendas"]
        )
        if tipo_consulta == "Una tienda":
            tienda_sel = st.sidebar.selectbox("Selecciona tienda", departamentos_disponibles)
            tiendas = [tienda_sel] if tienda_sel else []
        else:
            tiendas = st.sidebar.multiselect(
                "Selecciona tiendas",
                departamentos_disponibles,
                default=departamentos_disponibles,
            )
    
    if not tiendas or not meses_sel:
        st.warning("Selecciona al menos una tienda y un mes.")
        st.stop()
    
    nombre_meses_str = (
        ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} meses"
    )
    
    if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":
        st.subheader(f"Comparativa Interanual: {nombre_meses_str} ({ano} vs {ano - 1})")
        
        ano_anterior = ano - 1
        datos_fuente = {}
        
        df_ant = obtener_filtro_datos(df, ano_anterior, meses_sel, tiendas)
        df_act = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        
        datos_fuente["Ant"] = calcular_resultados(df_ant)
        datos_fuente["Act"] = calcular_resultados(df_act)
        
        columnas_eje = [f"Total {ano}", f"Total {ano_anterior}", "Var. €", "Var. %"]
        columnas_tabla = ["Resultados"] + columnas_eje
        filas_tabla_display = []
        filas_valores_numericos = []
        
        for concepto in CONCEPTOS_KPI:
            fila_disp = [concepto]
            fila_num = [concepto]
            
            val_ant = datos_fuente["Ant"].get(concepto, 0.0)
            val_act = datos_fuente["Act"].get(concepto, 0.0)
            
            if concepto == "R. B.":
                var_diff = val_act - val_ant
                fila_num.extend([val_act, val_ant, var_diff, 0.0])
                fila_disp.append(formato_porcentaje(val_act))
                fila_disp.append(formato_porcentaje(val_ant))
                fila_disp.append(formato_variacion_pp(var_diff))
                fila_disp.append("-")
            else:
                var_eur = val_act - val_ant
                var_pct = (var_eur / abs(val_ant) * 100) if val_ant != 0 else 0.0
                fila_num.extend([val_act, val_ant, var_eur, var_pct])
                fila_disp.append(formato_moneda(val_act))
                fila_disp.append(formato_moneda(val_ant))
                fila_disp.append(formato_moneda(var_eur))
                fila_disp.append(formato_porcentaje(var_pct / 100) if var_pct != 0 else "0,00%")
            
            filas_tabla_display.append(fila_disp)
            filas_valores_numericos.append(fila_num)
        
        df_resultado_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
        df_valores_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)
        
        render_aggrid_table(df_resultado_display, modo="auto")
        descargar_excel(df_valores_numericos, "Informe", f"Informe_Resultados_{ano}.xlsx",
                       "Descargar Informe en Excel")
    
    elif modo_analisis == "Comparativa Multi-Tienda (Totales)":
        st.subheader(f"Comparativa Multi-Tienda ({nombre_meses_str} {ano})")
        
        datos_fuente = {}
        for tienda in tiendas:
            df_filtrado = obtener_filtro_datos(df, ano, meses_sel, [tienda])
            datos_fuente[tienda] = calcular_resultados(df_filtrado)
        
        df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas)
        datos_fuente["Total"] = calcular_resultados(df_filtrado)
        
        columnas_eje = tiendas + ["Total"]
        columnas_tabla = ["Resultados"] + columnas_eje
        filas_tabla_display = []
        filas_valores_numericos = []
        
        for concepto in CONCEPTOS_KPI:
            fila_disp = [concepto]
            fila_num = [concepto]
            
            for item in columnas_eje:
                val = datos_fuente[item].get(concepto, 0.0)
                fila_num.append(val)
                
                if concepto == "R. B.":
                    fila_disp.append(formato_porcentaje(val))
                else:
                    fila_disp.append(formato_moneda(val))
            
            filas_tabla_display.append(fila_disp)
            filas_valores_numericos.append(fila_num)
        
        df_resultado_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
        df_valores_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)
        
        render_aggrid_table(df_resultado_display, modo="auto")
        descargar_excel(df_valores_numericos, "Informe", f"Informe_Resultados_{ano}.xlsx",
                       "Descargar Informe en Excel")
    
    else:  # Evolución Mensual / Tienda
        st.subheader(f"Informe ({nombre_meses_str} {ano})")
        
        datos_fuente = {}
        
        if len(tiendas) > 1:
            # Comparativa por tiendas
            for tienda in tiendas:
                df_filtrado = obtener_filtro_datos(df, ano, meses_sel, [tienda])
                datos_fuente[tienda] = calcular_resultados(df_filtrado)
            
            df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas)
            datos_fuente["Total"] = calcular_resultados(df_filtrado)
            columnas_eje = tiendas + ["Total"]
        else:
            # Evolución por meses
            for mes in meses_sel:
                df_filtrado = obtener_filtro_datos(df, ano, [mes], tiendas)
                datos_fuente[mes] = calcular_resultados(df_filtrado)
            
            if len(meses_sel) > 1:
                df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas)
                datos_fuente["Total"] = calcular_resultados(df_filtrado)
                columnas_eje = meses_sel + ["Total"]
            else:
                columnas_eje = meses_sel
        
        columnas_tabla = ["Resultados"] + columnas_eje
        filas_tabla_display = []
        filas_valores_numericos = []
        
        for concepto in CONCEPTOS_KPI:
            fila_disp = [concepto]
            fila_num = [concepto]
            
            for item in columnas_eje:
                val = datos_fuente[item].get(concepto, 0.0)
                fila_num.append(val)
                
                if concepto == "R. B.":
                    fila_disp.append(formato_porcentaje(val))
                else:
                    fila_disp.append(formato_moneda(val))
            
            filas_tabla_display.append(fila_disp)
            filas_valores_numericos.append(fila_num)
        
        df_resultado_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
        df_valores_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)
        
        render_aggrid_table(df_resultado_display, modo="auto")
        descargar_excel(df_valores_numericos, "Informe", f"Informe_Resultados_{ano}.xlsx",
                       "Descargar Informe en Excel")
