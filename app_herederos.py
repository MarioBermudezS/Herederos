import io
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

# AgGrid es opcional: si no está instalado, la aplicación usa st.dataframe.
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
    AGGRID_AVAILABLE = True
except ImportError:
    AGGRID_AVAILABLE = False


# ============================================================================
# CONFIGURACIÓN
# ============================================================================

st.set_page_config(
    page_title="Control de Resultados - Herederos",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "BaseDatos2026.xlsx"
SHEET_NAME = "BS"

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

RB_NORM = {"R. B.", "R.B.", "RB"}

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

h1 {font-size: 1.2rem !important; margin-bottom: 0.1rem !important;}
h2 {font-size: 1.05rem !important;}
h3 {font-size: 0.95rem !important; margin-bottom: 0.1rem !important;}
</style>
"""

st.markdown(CSS_ESTILOS, unsafe_allow_html=True)
st.title("Control de Resultados - Herederos")


# ============================================================================
# FORMATO
# ============================================================================

def formato_porcentaje(valor: float, decimales: int = 2) -> str:
    """Decimal -> porcentaje con formato español."""
    if pd.isna(valor):
        valor = 0.0
    return (
        f"{float(valor) * 100:,.{decimales}f}%"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def formato_moneda(valor: float, decimales: int = 2) -> str:
    """Número -> euros con formato español."""
    if pd.isna(valor):
        valor = 0.0
    return (
        f"{float(valor):,.{decimales}f} €"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def formato_variacion_pp(valor: float) -> str:
    """Diferencia de ratios expresada en puntos porcentuales."""
    return (
        f"{float(valor) * 100:,.2f} pp"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# ============================================================================
# CARGA Y NORMALIZACIÓN DE DATOS
# ============================================================================

def normalizar_resultado(valor) -> str:
    if pd.isna(valor):
        return ""
    return " ".join(str(valor).strip().split())


def normalizar_mes(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = " ".join(str(valor).strip().split())
    mapa = {m.lower(): m for m in MESES_ORDEN}
    return mapa.get(texto.lower(), texto)


@st.cache_data
def load_data() -> pd.DataFrame:
    """Lee y normaliza BaseDatos2026.xlsx / hoja BS."""
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"No se encuentra el archivo '{EXCEL_FILE.name}' "
            f"en la carpeta de la aplicación: {EXCEL_FILE.parent}"
        )

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    columnas_obligatorias = {
        "Año",
        "Mes",
        "Departamento",
        "Resultados",
        "Importe D",
    }
    faltan = columnas_obligatorias.difference(df.columns)
    if faltan:
        raise ValueError(
            "Faltan columnas obligatorias en la hoja BS: "
            + ", ".join(sorted(faltan))
        )

    # Limpieza de columnas.
    df = df.copy()
    df["Año"] = pd.to_numeric(df["Año"], errors="coerce").astype("Int64")
    df["Mes"] = df["Mes"].map(normalizar_mes)
    df["Departamento"] = df["Departamento"].fillna("").astype(str).str.strip()
    df["Resultados"] = df["Resultados"].map(normalizar_resultado)
    df["Importe D"] = pd.to_numeric(df["Importe D"], errors="coerce").fillna(0.0)

    # Eliminamos filas sin año/departamento/resultado.
    df = df[
        df["Año"].notna()
        & df["Departamento"].ne("")
        & df["Resultados"].ne("")
    ].copy()

    df["Año"] = df["Año"].astype(int)

    # Columna auxiliar para identificar R.B. sin depender de mayúsculas/espacios.
    df["_resultado_norm"] = (
        df["Resultados"]
        .str.upper()
        .str.replace(" ", "", regex=False)
    )

    return df


def obtener_filtro_datos(
    df: pd.DataFrame,
    ano: int,
    meses: List[str],
    departamentos: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Filtra por año, meses y, opcionalmente, departamentos."""
    if not meses:
        return df.iloc[0:0].copy()

    mascara = (df["Año"] == ano) & df["Mes"].isin(meses)

    if departamentos:
        mascara &= df["Departamento"].isin(departamentos)

    return df.loc[mascara].copy()


# ============================================================================
# CÁLCULOS
# ============================================================================

def _valor_concepto(df: pd.DataFrame, concepto: str) -> float:
    """Suma Importe D para un concepto."""
    if df.empty:
        return 0.0
    return float(df.loc[df["Resultados"] == concepto, "Importe D"].sum())


def _rb_por_grupo(group: pd.DataFrame) -> float:
    """
    Obtiene el R.B. del grupo Año-Mes-Departamento.

    El Excel contiene R.B. como decimal (ej. 0,4397 = 43,97%).
    Si hubiera más de una fila R.B. en el mismo grupo, se utiliza la
    primera y se genera una advertencia en el diagnóstico.
    """
    filas = group.loc[group["_resultado_norm"].isin(RB_NORM), "Importe D"]

    if filas.empty:
        return 0.0

    return float(filas.iloc[0])


def calcular_rb_puro(df_filtrado: pd.DataFrame) -> float:
    """
    Calcula R.B. ponderando el margen de cada Año-Mes-Departamento
    por sus ventas.

    Fórmula:
        R.B. grupo = Ventas grupo × R.B. grupo
        R.B. total = SUM(margen) / SUM(ventas)
    """
    if df_filtrado.empty:
        return 0.0

    total_ventas = 0.0
    total_margen = 0.0

    for _, group in df_filtrado.groupby(
        ["Año", "Mes", "Departamento"],
        dropna=False,
    ):
        ventas = float(
            group.loc[group["Resultados"] == "Ventas", "Importe D"].sum()
        )
        rb_val = _rb_por_grupo(group)

        total_ventas += ventas
        total_margen += ventas * rb_val

    if total_ventas == 0:
        # Mantiene el comportamiento del programa original como respaldo.
        rb_rows = df_filtrado.loc[
            df_filtrado["_resultado_norm"].isin(RB_NORM),
            "Importe D",
        ]
        return float(rb_rows.sum()) if not rb_rows.empty else 0.0

    return total_margen / total_ventas


def calcular_resultados(df_filtrado: pd.DataFrame) -> Dict[str, float]:
    """
    Calcula la cuenta de resultados.

    Importante:
    - R.B. se trata como porcentaje decimal almacenado en el Excel.
    - El coste de ventas se obtiene como Ventas - Margen Bruto.
    - Se mantiene la lógica contable del archivo original.
    """
    if df_filtrado.empty:
        return {concepto: 0.0 for concepto in CONCEPTOS_KPI}

    ventas = _valor_concepto(df_filtrado, "Ventas")

    # Margen bruto ponderado por ventas.
    total_ventas_calc = 0.0
    total_margen = 0.0

    for _, group in df_filtrado.groupby(
        ["Año", "Mes", "Departamento"],
        dropna=False,
    ):
        ventas_grupo = float(
            group.loc[group["Resultados"] == "Ventas", "Importe D"].sum()
        )
        rb_val = _rb_por_grupo(group)

        total_ventas_calc += ventas_grupo
        total_margen += ventas_grupo * rb_val

    if total_ventas_calc != 0:
        margen_bruto = total_margen
        r_bruta = margen_bruto / total_ventas_calc
    else:
        # Respaldo para conjuntos sin ventas calculables.
        rb_rows = df_filtrado.loc[
            df_filtrado["_resultado_norm"].isin(RB_NORM),
            "Importe D",
        ]
        r_bruta = float(rb_rows.sum()) if not rb_rows.empty else 0.0
        margen_bruto = r_bruta * ventas

    coste_ventas = ventas - margen_bruto

    # Ingresos operativos.
    otros_ingresos = _valor_concepto(df_filtrado, "Otros Ingresos")
    ingresos_operativos = margen_bruto + otros_ingresos

    # Gastos operativos.
    gastos_personal = _valor_concepto(df_filtrado, "Gastos Personal")
    alquileres = _valor_concepto(df_filtrado, "Alquileres")
    reparaciones = _valor_concepto(df_filtrado, "Reparaciones")
    seguros = _valor_concepto(df_filtrado, "Seguros")
    suministros = _valor_concepto(df_filtrado, "Suministros")
    otros_servicios = _valor_concepto(df_filtrado, "Otros Servicios")

    total_gastos_operativos = (
        alquileres
        + reparaciones
        + seguros
        + suministros
        + otros_servicios
    )

    amortizaciones = _valor_concepto(df_filtrado, "Amortizaciones")

    gastos_estructura = (
        total_gastos_operativos
        + gastos_personal
        + amortizaciones
    )

    baii = (
        ingresos_operativos
        - gastos_personal
        - total_gastos_operativos
        - amortizaciones
    )

    # Resultado financiero.
    gastos_financieros = _valor_concepto(df_filtrado, "Gastos Financieros")
    ingresos_financieros = _valor_concepto(df_filtrado, "Ingresos Financieros")

    # En la base, los ingresos financieros están almacenados con signo
    # negativo; por ello se mantiene la suma del original.
    rdo_financiero = gastos_financieros + ingresos_financieros

    resultados_extraordinarios = _valor_concepto(
        df_filtrado,
        "Resultados Extraordinarios",
    )

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


# ============================================================================
# TABLAS / EXPORTACIÓN
# ============================================================================

def calcular_ancho_columna(
    df: pd.DataFrame,
    col_name: str,
    min_width: int = 80,
) -> int:
    if col_name not in df.columns:
        return min_width

    try:
        max_len = max(
            len(str(v)) for v in df[col_name].astype(str)
        )
    except Exception:
        max_len = len(str(col_name))

    ancho = max(
        max_len * 8 + 15,
        len(str(col_name)) * 8 + 15,
        min_width,
    )
    return min(ancho, 250)


def render_aggrid_table(
    df_display: pd.DataFrame,
    modo: str = "auto",
) -> None:
    """Muestra la tabla con AgGrid o, como respaldo, st.dataframe."""
    if df_display.empty:
        st.info("No hay datos para mostrar.")
        return

    if not AGGRID_AVAILABLE:
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
        )
        return

    gb = GridOptionsBuilder.from_dataframe(df_display)

    gb.configure_default_column(
        resizable=True,
        filterable=False,
        sortable=False,
        editable=False,
        suppressMenu=True,
    )

    if len(df_display.columns) > 0:
        first_col = df_display.columns[0]
        gb.configure_column(
            first_col,
            pinned="left",
            width=calcular_ancho_columna(df_display, first_col, 200),
            minWidth=150,
            cellStyle={"fontWeight": "bold", "textAlign": "left"},
        )

    for col in df_display.columns[1:]:
        gb.configure_column(
            col,
            width=calcular_ancho_columna(df_display, col, 100),
            minWidth=80,
            cellStyle={"textAlign": "right"},
        )

    gb.configure_grid_options(
        domLayout="autoHeight" if modo == "auto" else "normal",
        suppressRowClickSelection=True,
    )

    AgGrid(
        df_display,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=True,
        allow_unsafe_jscode=True,
        theme="balham",
        height=400,
    )


def descargar_excel(
    df: pd.DataFrame,
    nombre_hoja: str,
    nombre_archivo: str,
    etiqueta: str,
) -> None:
    """Genera un Excel descargable."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name=nombre_hoja[:31],
        )

    st.download_button(
        label=f"📥 {etiqueta}",
        data=output.getvalue(),
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# ============================================================================
# TABLAS DE RESULTADOS
# ============================================================================

def construir_tabla_kpi(
    datos_fuente: Dict[str, Dict[str, float]],
    columnas_eje: List[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    columnas = ["Resultados"] + columnas_eje
    display_rows = []
    numeric_rows = []

    for concepto in CONCEPTOS_KPI:
        fila_display = [concepto]
        fila_numeric = [concepto]

        for item in columnas_eje:
            resultados = datos_fuente[item]
            ventas = resultados.get("Ventas", 0.0)
            valor = resultados.get(concepto, 0.0)

            if concepto == "R. B.":
                kpi = valor
            else:
                kpi = valor / ventas if ventas != 0 else 0.0

            fila_numeric.append(kpi)
            fila_display.append(formato_porcentaje(kpi))

        display_rows.append(fila_display)
        numeric_rows.append(fila_numeric)

    return (
        pd.DataFrame(display_rows, columns=columnas),
        pd.DataFrame(numeric_rows, columns=columnas),
    )


def construir_tabla_resultados(
    datos_fuente: Dict[str, Dict[str, float]],
    columnas_eje: List[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    columnas = ["Resultados"] + columnas_eje
    display_rows = []
    numeric_rows = []

    for concepto in CONCEPTOS_KPI:
        fila_display = [concepto]
        fila_numeric = [concepto]

        for item in columnas_eje:
            valor = datos_fuente[item].get(concepto, 0.0)
            fila_numeric.append(valor)

            if concepto == "R. B.":
                fila_display.append(formato_porcentaje(valor))
            else:
                fila_display.append(formato_moneda(valor))

        display_rows.append(fila_display)
        numeric_rows.append(fila_numeric)

    return (
        pd.DataFrame(display_rows, columns=columnas),
        pd.DataFrame(numeric_rows, columns=columnas),
    )


# ============================================================================
# CARGA
# ============================================================================

try:
    df = load_data()
except Exception as exc:
    st.error(f"Error al leer '{EXCEL_FILE.name}': {exc}")
    st.stop()


# ============================================================================
# SIDEBAR GENERAL
# ============================================================================

st.sidebar.header("Parámetros del Informe")

modulo_principal = st.sidebar.radio(
    "Módulo de Análisis",
    [
        "Cuenta de Resultados Completa",
        "Análisis Específico de R.B. (Margen Bruto)",
        "Informe KPI (% sobre Ventas)",
    ],
)

anos_disponibles = sorted(df["Año"].unique().tolist())

if not anos_disponibles:
    st.error("No hay años válidos en la hoja BS.")
    st.stop()

ano = st.sidebar.selectbox(
    "Año principal",
    anos_disponibles,
    index=len(anos_disponibles) - 1,
)

meses_excel = df.loc[df["Año"] == ano, "Mes"].dropna().unique().tolist()
meses_disponibles = [m for m in MESES_ORDEN if m in meses_excel]

# Por seguridad, añadimos meses no estándar al final.
meses_disponibles += [
    m for m in meses_excel if m not in meses_disponibles
]

if not meses_disponibles:
    st.error(f"No hay meses disponibles para {ano}.")
    st.stop()

meses_sel = st.sidebar.multiselect(
    "Selecciona mes(es)",
    meses_disponibles,
    default=meses_disponibles[:1],
)

df_ano = df.loc[df["Año"] == ano]

departamentos_disponibles = sorted(
    df_ano["Departamento"].dropna().unique().tolist()
)


# ============================================================================
# MÓDULO 1 - R.B.
# ============================================================================

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

    if tipo_analisis_rb == "Evolución Mensual por Tienda":

        st.subheader(
            f"Análisis R.B. - Evolución Mensual por Tienda ({ano})"
        )

        columnas = ["Resultados"] + meses_sel
        if len(meses_sel) > 1:
            columnas.append("Promedio Acumulado")

        display_rows = []
        numeric_rows = []

        for tienda in tiendas_rb:
            display = [tienda]
            numeric = [tienda]

            for mes in meses_sel:
                datos = obtener_filtro_datos(
                    df, ano, [mes], [tienda]
                )
                rb = calcular_rb_puro(datos)

                numeric.append(rb)
                display.append(formato_porcentaje(rb))

            if len(meses_sel) > 1:
                datos = obtener_filtro_datos(
                    df, ano, meses_sel, [tienda]
                )
                rb = calcular_rb_puro(datos)

                numeric.append(rb)
                display.append(formato_porcentaje(rb))

            display_rows.append(display)
            numeric_rows.append(numeric)

        if len(tiendas_rb) > 1:
            display = ["TOTAL GRUPO"]
            numeric = ["TOTAL GRUPO"]

            for mes in meses_sel:
                datos = obtener_filtro_datos(
                    df, ano, [mes], tiendas_rb
                )
                rb = calcular_rb_puro(datos)

                numeric.append(rb)
                display.append(formato_porcentaje(rb))

            if len(meses_sel) > 1:
                datos = obtener_filtro_datos(
                    df, ano, meses_sel, tiendas_rb
                )
                rb = calcular_rb_puro(datos)

                numeric.append(rb)
                display.append(formato_porcentaje(rb))

            display_rows.append(display)
            numeric_rows.append(numeric)

        df_display = pd.DataFrame(display_rows, columns=columnas)
        df_numeric = pd.DataFrame(numeric_rows, columns=columnas)

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Analisis_RB_Mensual",
            f"Analisis_RB_Mensual_{ano}.xlsx",
            "Descargar Análisis R.B. en Excel",
        )

    elif tipo_analisis_rb == "Vista Acumulada por Tienda":

        nombre_meses = (
            ", ".join(meses_sel)
            if len(meses_sel) <= 3
            else f"{len(meses_sel)} meses acumulados"
        )

        st.subheader(
            f"Análisis R.B. - Vista Acumulada ({nombre_meses} {ano})"
        )

        display_rows = []
        numeric_rows = []

        for tienda in tiendas_rb:
            datos = obtener_filtro_datos(
                df, ano, meses_sel, [tienda]
            )
            rb = calcular_rb_puro(datos)

            display_rows.append(
                [tienda, formato_porcentaje(rb)]
            )
            numeric_rows.append([tienda, rb])

        if len(tiendas_rb) > 1:
            datos = obtener_filtro_datos(
                df, ano, meses_sel, tiendas_rb
            )
            rb = calcular_rb_puro(datos)

            display_rows.append(
                ["TOTAL GRUPO", formato_porcentaje(rb)]
            )
            numeric_rows.append(["TOTAL GRUPO", rb])

        nombre_col = f"Acumulado {nombre_meses}"

        df_display = pd.DataFrame(
            display_rows,
            columns=["Resultados", nombre_col],
        )
        df_numeric = pd.DataFrame(
            numeric_rows,
            columns=["Resultados", nombre_col],
        )

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Analisis_RB_Acumulado",
            f"Analisis_RB_Acumulado_{ano}.xlsx",
            "Descargar Acumulado R.B. en Excel",
        )

    else:

        ano_ant = ano - 1
        nombre_meses = (
            ", ".join(meses_sel)
            if len(meses_sel) <= 3
            else f"{len(meses_sel)} meses"
        )

        st.subheader(
            f"Comparativa Interanual R.B. ({nombre_meses}): "
            f"{ano} vs {ano_ant}"
        )

        columnas = [
            "Resultados",
            f"R.B. {ano}",
            f"R.B. {ano_ant}",
            "Var. pp",
        ]

        display_rows = []
        numeric_rows = []

        for tienda in tiendas_rb:
            actual = obtener_filtro_datos(
                df, ano, meses_sel, [tienda]
            )
            anterior = obtener_filtro_datos(
                df, ano_ant, meses_sel, [tienda]
            )

            rb_act = calcular_rb_puro(actual)
            rb_ant = calcular_rb_puro(anterior)
            variacion = rb_act - rb_ant

            numeric_rows.append(
                [tienda, rb_act, rb_ant, variacion]
            )
            display_rows.append(
                [
                    tienda,
                    formato_porcentaje(rb_act),
                    formato_porcentaje(rb_ant),
                    formato_variacion_pp(variacion),
                ]
            )

        if len(tiendas_rb) > 1:
            actual = obtener_filtro_datos(
                df, ano, meses_sel, tiendas_rb
            )
            anterior = obtener_filtro_datos(
                df, ano_ant, meses_sel, tiendas_rb
            )

            rb_act = calcular_rb_puro(actual)
            rb_ant = calcular_rb_puro(anterior)
            variacion = rb_act - rb_ant

            numeric_rows.append(
                ["TOTAL GRUPO", rb_act, rb_ant, variacion]
            )
            display_rows.append(
                [
                    "TOTAL GRUPO",
                    formato_porcentaje(rb_act),
                    formato_porcentaje(rb_ant),
                    formato_variacion_pp(variacion),
                ]
            )

        df_display = pd.DataFrame(display_rows, columns=columnas)
        df_numeric = pd.DataFrame(numeric_rows, columns=columnas)

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Interanual_RB",
            f"Comparativa_Interanual_RB_{ano}.xlsx",
            "Descargar Comparativa R.B. en Excel",
        )


# ============================================================================
# MÓDULO 2 - KPI
# ============================================================================

elif modulo_principal == "Informe KPI (% sobre Ventas)":

    modo_analisis = st.sidebar.radio(
        "Tipo de Análisis KPI",
        [
            "Evolución Mensual / Tienda",
            "Comparativa Multi-Tienda (Totales)",
            "Comparativa Interanual (Año vs Año Anterior)",
        ],
    )

    if modo_analisis == "Comparativa Multi-Tienda (Totales)":
        tiendas = st.sidebar.multiselect(
            "Selecciona tiendas a comparar",
            departamentos_disponibles,
            default=(
                departamentos_disponibles[:2]
                if len(departamentos_disponibles) >= 2
                else departamentos_disponibles
            ),
        )
    else:
        tipo_consulta = st.sidebar.radio(
            "Tipo de consulta",
            ["Una tienda", "Conjunto de tiendas"],
        )

        if tipo_consulta == "Una tienda":
            tienda_sel = st.sidebar.selectbox(
                "Selecciona tienda",
                departamentos_disponibles,
            )
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

    nombre_meses = (
        ", ".join(meses_sel)
        if len(meses_sel) <= 3
        else f"{len(meses_sel)} meses"
    )

    if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":

        ano_ant = ano - 1

        st.subheader(
            f"Informe KPI (% sobre Ventas) - Interanual: "
            f"{nombre_meses} ({ano} vs {ano_ant})"
        )

        datos_fuente = {
            "Act": calcular_resultados(
                obtener_filtro_datos(df, ano, meses_sel, tiendas)
            ),
            "Ant": calcular_resultados(
                obtener_filtro_datos(df, ano_ant, meses_sel, tiendas)
            ),
        }

        columnas = [
            "Resultados",
            f"Total {ano}",
            f"Total {ano_ant}",
            "Var. pp",
        ]

        display_rows = []
        numeric_rows = []

        for concepto in CONCEPTOS_KPI:
            act = datos_fuente["Act"]
            ant = datos_fuente["Ant"]

            ventas_act = act.get("Ventas", 0.0)
            ventas_ant = ant.get("Ventas", 0.0)

            valor_act = act.get(concepto, 0.0)
            valor_ant = ant.get(concepto, 0.0)

            if concepto == "R. B.":
                kpi_act = valor_act
                kpi_ant = valor_ant
            else:
                kpi_act = (
                    valor_act / ventas_act
                    if ventas_act != 0
                    else 0.0
                )
                kpi_ant = (
                    valor_ant / ventas_ant
                    if ventas_ant != 0
                    else 0.0
                )

            var_pp = kpi_act - kpi_ant

            numeric_rows.append(
                [concepto, kpi_act, kpi_ant, var_pp]
            )
            display_rows.append(
                [
                    concepto,
                    formato_porcentaje(kpi_act),
                    formato_porcentaje(kpi_ant),
                    formato_variacion_pp(var_pp),
                ]
            )

        df_display = pd.DataFrame(display_rows, columns=columnas)
        df_numeric = pd.DataFrame(numeric_rows, columns=columnas)

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Informe_KPI",
            f"Informe_KPI_Ventas_{ano}.xlsx",
            "Descargar Informe KPI en Excel",
        )

    elif modo_analisis == "Comparativa Multi-Tienda (Totales)":

        st.subheader(
            f"Informe KPI (% sobre Ventas) - Multi-Tienda "
            f"({nombre_meses} {ano})"
        )

        datos_fuente = {}

        for tienda in tiendas:
            datos_fuente[tienda] = calcular_resultados(
                obtener_filtro_datos(df, ano, meses_sel, [tienda])
            )

        datos_fuente["Total"] = calcular_resultados(
            obtener_filtro_datos(df, ano, meses_sel, tiendas)
        )

        columnas_eje = tiendas + ["Total"]
        df_display, df_numeric = construir_tabla_kpi(
            datos_fuente,
            columnas_eje,
        )

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Informe_KPI",
            f"Informe_KPI_Ventas_{ano}.xlsx",
            "Descargar Informe KPI en Excel",
        )

    else:

        st.subheader(
            f"Informe KPI (% sobre Ventas) - "
            f"({nombre_meses} {ano})"
        )

        datos_fuente = {}

        if len(tiendas) > 1:
            for tienda in tiendas:
                datos_fuente[tienda] = calcular_resultados(
                    obtener_filtro_datos(
                        df, ano, meses_sel, [tienda]
                    )
                )

            datos_fuente["Total"] = calcular_resultados(
                obtener_filtro_datos(
                    df, ano, meses_sel, tiendas
                )
            )

            columnas_eje = tiendas + ["Total"]

        else:
            for mes in meses_sel:
                datos_fuente[mes] = calcular_resultados(
                    obtener_filtro_datos(
                        df, ano, [mes], tiendas
                    )
                )

            if len(meses_sel) > 1:
                datos_fuente["Total"] = calcular_resultados(
                    obtener_filtro_datos(
                        df, ano, meses_sel, tiendas
                    )
                )
                columnas_eje = meses_sel + ["Total"]
            else:
                columnas_eje = meses_sel

        df_display, df_numeric = construir_tabla_kpi(
            datos_fuente,
            columnas_eje,
        )

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Informe_KPI",
            f"Informe_KPI_Ventas_{ano}.xlsx",
            "Descargar Informe KPI en Excel",
        )


# ============================================================================
# MÓDULO 3 - CUENTA DE RESULTADOS
# ============================================================================

else:

    modo_analisis = st.sidebar.radio(
        "Tipo de Análisis",
        [
            "Evolución Mensual / Tienda",
            "Comparativa Multi-Tienda (Totales)",
            "Comparativa Interanual (Año vs Año Anterior)",
        ],
    )

    if modo_analisis == "Comparativa Multi-Tienda (Totales)":
        tiendas = st.sidebar.multiselect(
            "Selecciona tiendas a comparar",
            departamentos_disponibles,
            default=(
                departamentos_disponibles[:2]
                if len(departamentos_disponibles) >= 2
                else departamentos_disponibles
            ),
        )
    else:
        tipo_consulta = st.sidebar.radio(
            "Tipo de consulta",
            ["Una tienda", "Conjunto de tiendas"],
        )

        if tipo_consulta == "Una tienda":
            tienda_sel = st.sidebar.selectbox(
                "Selecciona tienda",
                departamentos_disponibles,
            )
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

    nombre_meses = (
        ", ".join(meses_sel)
        if len(meses_sel) <= 3
        else f"{len(meses_sel)} meses"
    )

    if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":

        ano_ant = ano - 1

        st.subheader(
            f"Comparativa Interanual: "
            f"{nombre_meses} ({ano} vs {ano_ant})"
        )

        datos_fuente = {
            "Act": calcular_resultados(
                obtener_filtro_datos(df, ano, meses_sel, tiendas)
            ),
            "Ant": calcular_resultados(
                obtener_filtro_datos(df, ano_ant, meses_sel, tiendas)
            ),
        }

        columnas = [
            "Resultados",
            f"Total {ano}",
            f"Total {ano_ant}",
            "Var. €",
            "Var. %",
        ]

        display_rows = []
        numeric_rows = []

        for concepto in CONCEPTOS_KPI:

            val_act = datos_fuente["Act"].get(concepto, 0.0)
            val_ant = datos_fuente["Ant"].get(concepto, 0.0)

            if concepto == "R. B.":

                var = val_act - val_ant

                numeric_rows.append(
                    [concepto, val_act, val_ant, var, 0.0]
                )
                display_rows.append(
                    [
                        concepto,
                        formato_porcentaje(val_act),
                        formato_porcentaje(val_ant),
                        formato_variacion_pp(var),
                        "-",
                    ]
                )

            else:

                var_eur = val_act - val_ant
                var_pct = (
                    var_eur / abs(val_ant)
                    if val_ant != 0
                    else 0.0
                )

                numeric_rows.append(
                    [concepto, val_act, val_ant, var_eur, var_pct]
                )
                display_rows.append(
                    [
                        concepto,
                        formato_moneda(val_act),
                        formato_moneda(val_ant),
                        formato_moneda(var_eur),
                        formato_porcentaje(var_pct),
                    ]
                )

        df_display = pd.DataFrame(display_rows, columns=columnas)
        df_numeric = pd.DataFrame(numeric_rows, columns=columnas)

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Informe",
            f"Informe_Resultados_{ano}.xlsx",
            "Descargar Informe en Excel",
        )

    elif modo_analisis == "Comparativa Multi-Tienda (Totales)":

        st.subheader(
            f"Comparativa Multi-Tienda ({nombre_meses} {ano})"
        )

        datos_fuente = {}

        for tienda in tiendas:
            datos_fuente[tienda] = calcular_resultados(
                obtener_filtro_datos(df, ano, meses_sel, [tienda])
            )

        datos_fuente["Total"] = calcular_resultados(
            obtener_filtro_datos(df, ano, meses_sel, tiendas)
        )

        columnas_eje = tiendas + ["Total"]

        df_display, df_numeric = construir_tabla_resultados(
            datos_fuente,
            columnas_eje,
        )

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Informe",
            f"Informe_Resultados_{ano}.xlsx",
            "Descargar Informe en Excel",
        )

    else:

        st.subheader(
            f"Informe ({nombre_meses} {ano})"
        )

        datos_fuente = {}

        if len(tiendas) > 1:

            for tienda in tiendas:
                datos_fuente[tienda] = calcular_resultados(
                    obtener_filtro_datos(
                        df, ano, meses_sel, [tienda]
                    )
                )

            datos_fuente["Total"] = calcular_resultados(
                obtener_filtro_datos(
                    df, ano, meses_sel, tiendas
                )
            )

            columnas_eje = tiendas + ["Total"]

        else:

            for mes in meses_sel:
                datos_fuente[mes] = calcular_resultados(
                    obtener_filtro_datos(
                        df, ano, [mes], tiendas
                    )
                )

            if len(meses_sel) > 1:
                datos_fuente["Total"] = calcular_resultados(
                    obtener_filtro_datos(
                        df, ano, meses_sel, tiendas
                    )
                )
                columnas_eje = meses_sel + ["Total"]
            else:
                columnas_eje = meses_sel

        df_display, df_numeric = construir_tabla_resultados(
            datos_fuente,
            columnas_eje,
        )

        render_aggrid_table(df_display)
        descargar_excel(
            df_numeric,
            "Informe",
            f"Informe_Resultados_{ano}.xlsx",
            "Descargar Informe en Excel",
        )


# ============================================================================
# PIE
# ============================================================================

if not AGGRID_AVAILABLE:
    st.sidebar.warning(
        "AgGrid no está instalado. Se está usando la tabla nativa de "
        "Streamlit. Instala 'streamlit-aggrid' para recuperar las tablas "
        "interactivas."
    )
