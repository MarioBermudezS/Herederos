import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

# CSS ultra-compacto y forzoso para alinear a la derecha y evitar scroll
st.markdown(
    """
    <style>
    /* Expandir la ventana al máximo y eliminar padding */
    .block-container {
        padding-top: 0.3rem !important;
        padding-bottom: 0.3rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    h1 {
        font-size: 1.3rem !important;
        margin-bottom: 0.1rem !important;
    }
    h3 {
        font-size: 1.0rem !important;
        margin-bottom: 0.1rem !important;
    }
    /* Estilo de tabla hipercompacta estilo Excel para que quepa todo de un vistazo */
    table {
        width: 100% !important;
        font-size: 11.5px !important;
        border-collapse: collapse !important;
    }
    th, td {
        padding: 2px 5px !important;
        white-space: nowrap !important;
    }
    /* FORZAR ALINEACIÓN ABSOLUTA A LA DERECHA EN TODAS LAS COLUMNAS NUMÉRICAS Y CABECERAS */
    th:not(:first-child), td:not(:first-child) {
        text-align: right !important;
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

# Construcción de la tabla final
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

filas_tabla_display = []
filas_valores_numericos = []

for concepto in conceptos:
  fila_disp = [concepto]
  fila_num = [concepto]

  for mes in meses_sel:
    val = datos_por_mes[mes].get(concepto, 0.0)
    fila_num.append(val)
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

  if len(meses_sel) > 1:
    if concepto == "R. B.":
      tot_ventas = sum(datos_por_mes[m].get("Ventas", 0.0) for m in meses_sel)
      tot_margen = sum(
          datos_por_mes[m].get("MARGEN BRUTO", 0.0) for m in meses_sel
      )
      val_total = (tot_margen / tot_ventas) if tot_ventas != 0 else 0.0
      fila_num.append(val_total)
      fila_disp.append(
          f"{val_total * 100:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
    else:
      val_total = sum(datos_por_mes[m].get(concepto, 0.0) for m in meses_sel)
      fila_num.append(val_total)
      fila_disp.append(
          f"{val_total:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )

  filas_tabla_display.append(fila_disp)
  filas_valores_numericos.append(fila_num)

df_resultado_display = pd.DataFrame(
    filas_tabla_display, columns=columnas_tabla
)
df_valores_numericos = pd.DataFrame(
    filas_valores_numericos, columns=columnas_tabla
)

# Campos que deben ir destacados en negrita
campos_destacados = [
    "MARGEN BRUTO",
    "R. B.",
    "Ingresos Operativos",
    "GASTOS ESTRUCTURA",
    "B.A.I.I.",
    "RDO. FINANCIERO",
    "B.A.I.",
]


# Función de estilos Styler combinada con st.table para garantizar alineaciones y colores
def aplicar_estilos_styler(s):
  styles = []
  for i, row in df_resultado_display.iterrows():
    concepto = row["Resultados"]
    is_destacado = concepto in campos_destacados

    row_styles = [
        "text-align: left !important; padding-left: 6px;"
        + ("font-weight: bold; background-color: #eef2f7;" if is_destacado else "")
    ]

    for col_idx, col in enumerate(columnas_tabla[1:], start=1):
      num_val = df_valores_numericos.loc[i, col]
      is_negativo = isinstance(num_val, (int, float)) and num_val < 0
      is_columna_total = col == "Total"

      cell_style = "text-align: right !important; padding-right: 8px;"

      if is_columna_total:
        cell_style += " background-color: #d1fae5;"
      elif is_destacado:
        cell_style += " background-color: #eef2f7;"

      if is_destacado:
        cell_style += " font-weight: bold;"

      if is_negativo:
        cell_style += " color: #dc2626; font-weight: bold;"
      elif is_destacado:
        cell_style += " color: #1f2937;"

      row_styles.append(cell_style)
    styles.append(row_styles)

  return pd.DataFrame(styles, index=s.index, columns=s.columns)


df_styled = df_resultado_display.style.apply(aplicar_estilos_styler, axis=None)

# Usamos st.table para que pinte todas las filas de golpe de forma estática y compacta sin barras de scroll
st.table(df_styled)


# Botón de descarga directa en Excel (valores numéricos puros)
def to_excel(df_to_save):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_to_save.to_excel(writer, index=False, sheet_name="Informe")
  return output.getvalue()


excel_data = to_excel(df_valores_numericos)
nombre_salida = f"Informe_{'_'.join(tiendas)}_{ano}.xlsx"

st.download_button(
    label="📥 Descargar Informe en Excel",
    data=excel_data,
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
