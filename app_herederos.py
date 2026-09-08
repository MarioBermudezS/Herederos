import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

# CSS avanzado para asegurar que todas las celdas y cabeceras numéricas estén perfectamente alineadas a la derecha
st.markdown(
    """
    <style>
    /* Forzar alineación a la derecha en todas las columnas de datos excepto la primera (Resultados) */
    div[data-testid="stDataFrame"] td:not(:first-child), 
    div[data-testid="stDataFrame"] th:not(:first-child) {
        text-align: right !important;
        justify-content: flex-end !important;
    }
    div[data-testid="stDataFrame"] td:not(:first-child) div,
    div[data-testid="stDataFrame"] th:not(:first-child) div {
        text-align: right !important;
        justify-content: flex-end !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("Control de Resultados - Herederos")


@st.cache_data
def load_data():
  return pd.read_excel("BaseDatos2026.xlsx", sheet_name="BS")


try:
  df = load_data()
except Exception as e:
  st.error(
      f"Error al leer el archivo Excel ('BaseDatos2026.xlsx'): {e}"
  )
  st.stop()

# Menú lateral para filtros
st.sidebar.header("Parámetros del Informe")

# Restringir años exclusivamente de 2024 a 2026
anos_disponibles = [2024, 2025, 2026]
if "Año" in df.columns:
  anos_excel = sorted(df["Año"].dropna().unique())
  anos_disponibles = [a for a in anos_disponibles if a in anos_excel]
if not anos_disponibles:
  anos_disponibles = [2026]

ano = st.sidebar.selectbox("Año", anos_disponibles)

# Selección múltiple de meses ordenados cronológicamente
meses_orden = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
]
meses_excel = (
    df["Mes"].dropna().unique().tolist() if "Mes" in df.columns else ["Enero"]
)
meses_disponibles = [m for m in meses_orden if m in meses_excel]
if not meses_disponibles:
  meses_disponibles = meses_excel

meses_sel = st.sidebar.multiselect(
    "Selecciona mes(es)", meses_disponibles, default=meses_disponibles[:1]
)

departamentos_disponibles = (
    sorted(df["Departamento"].dropna().unique())
    if "Departamento" in df.columns
    else []
)

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
  st.warning(
      "Por favor, selecciona al menos una tienda y un mes en la barra lateral."
  )
  st.stop()

st.subheader(f"Informe para: {', '.join(tiendas)} ({ano})")


# Función para calcular los resultados de un conjunto de datos filtrados
def calcular_resultados(df_filtered):
  resumen = df_filtered.groupby("Resultados")["Importe D"].sum().to_dict()

  def get_v(cat):
    return resumen.get(cat, 0.0)

  ventas = get_v("Ventas")
  r_bruta = get_v("R. B.")
  if r_bruta == 0.0:
    for k, v in resumen.items():
      if k and str(k).strip().upper() in ["R. B.", "R.B.", "R.B"]:
        r_bruta = v
        break

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

  total_gastos_operativos = (
      alquileres + reparaciones + seguros + suministros + otros_servicios
  )
  amortizaciones = get_v("Amortizaciones")
  gastos_estructura = total_gastos_operativos + gastos_personal + amortizaciones

  baii = (
      ingresos_operativos
      - gastos_personal
      - total_gastos_operativos
      - amortizaciones
  )
  gastos_financieros = get_v("Gastos Financieros")
  ingresos_financieros = get_v("Ingresos Financieros")
  rdo_financiero = ingresos_financieros - gastos_financieros
  resultados_extraordinarios = get_v("Resultados Extraordinarios")
  bai = baii + rdo_financiero + resultados_extraordinarios

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


# Diccionario para almacenar los resultados de cada mes
datos_por_mes = {}
for mes in meses_sel:
  mask = (
      (df["Año"] == ano)
      & (df["Mes"] == mes)
      & (df["Departamento"].isin(tiendas))
  )
  df_m = df[mask]
  datos_por_mes[mes] = calcular_resultados(df_m)

# Construcción de la tabla final con columnas por mes + Total
conceptos = [
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

columnas_tabla = ["Resultados"] + meses_sel
if len(meses_sel) > 1:
  columnas_tabla.append("Total")

filas_tabla_raw = []
filas_tabla_display = []

for concepto in conceptos:
  fila_raw = [concepto]
  fila_disp = [concepto]

  for mes in meses_sel:
    val = datos_por_mes[mes].get(concepto, 0.0)
    fila_raw.append(val)
    if concepto == "R. B.":
      fila_disp.append(
          f"{val * 100:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
    else:
      fila_disp.append(
          f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
      )

  # Si hay más de un mes, calculamos la columna Total
  if len(meses_sel) > 1:
    if concepto == "R. B.":
      tot_ventas = sum(datos_por_mes[m].get("Ventas", 0.0) for m in meses_sel)
      tot_margen = sum(
          datos_por_mes[m].get("MARGEN BRUTO", 0.0) for m in meses_sel
      )
      val_total = (tot_margen / tot_ventas) if tot_ventas != 0 else 0.0
      fila_raw.append(val_total)
      fila_disp.append(
          f"{val_total * 100:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
    else:
      val_total = sum(datos_por_mes[m].get(concepto, 0.0) for m in meses_sel)
      fila_raw.append(val_total)
      fila_disp.append(
          f"{val_total:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )

  filas_tabla_raw.append(fila_raw)
  filas_tabla_display.append(fila_disp)

df_resultado_raw = pd.DataFrame(filas_tabla_raw, columns=columnas_tabla)
df_resultado_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)

# Campos que deben ir destacados en negrita y con fondo sutil
campos_destacados = [
    "MARGEN BRUTO",
    "R. B.",
    "Ingresos Operativos",
    "GASTOS ESTRUCTURA",
    "B.A.I.I.",
    "RDO. FINANCIERO",
    "B.A.I.",
]

# Estilizado visual con Pandas Styler
styles = []
for idx, row in df_resultado_display.iterrows():
  if row["Resultados"] in campos_destacados:
    styles.append(
        dict(
            selector=f"tr:nth-child({idx + 1})",
            props=[
                ("font-weight", "bold"),
                ("background-color", "#eef2f7"),
                ("color", "#1f2937"),
            ],
        )
    )

df_styled = df_resultado_display.style.set_table_styles(styles).set_properties(
    subset=["Resultados"], **{"text-align": "left", "padding-left": "10px"}
)

for col in columnas_tabla[1:]:
  df_styled = df_styled.set_properties(
      subset=[col], **{"text-align": "right", "padding-right": "15px"}
  )

# Mostrar tabla en pantalla
st.dataframe(df_styled, use_container_width=True, hide_index=True)


# Botón de descarga directa en Excel (valores numéricos puros)
def to_excel(df_to_save):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_to_save.to_excel(writer, index=False, sheet_name="Informe")
  return output.getvalue()


excel_data = to_excel(df_resultado_raw)
nombre_salida = f"Informe_{'_'.join(tiendas)}_{ano}.xlsx"

st.download_button(
    label="📥 Descargar Informe en Excel",
    data=excel_data,
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
