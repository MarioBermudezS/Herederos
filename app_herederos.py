import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

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

# Selección múltiple de meses
meses_disponibles = (
    df["Mes"].dropna().unique().tolist() if "Mes" in df.columns else ["Enero"]
)
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

nombre_meses_str = (
    ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} meses"
)
st.subheader(f"Informe para: {', '.join(tiendas)} ({nombre_meses_str} {ano})")

# Filtrado de datos por Año, Meses seleccionados y Tiendas
mask = (
    (df["Año"] == ano)
    & (df["Mes"].isin(meses_sel))
    & (df["Departamento"].isin(tiendas))
)
df_filtered = df[mask]

if df_filtered.empty:
  st.warning(
      "¡Aviso! No se han encontrado datos para esos criterios de selección."
  )
  st.stop()

resumen = df_filtered.groupby("Resultados")["Importe D"].sum().to_dict()


def get_val(cat):
  return resumen.get(cat, 0.0)


# Obtenemos Ventas y el valor de R. B. (si hay varios meses seleccionados, se promedia o acumula según criterio, aquí tomamos el valor medio de R.B. o directo)
ventas = get_val("Ventas")
r_bruta = get_val("R. B.")

if r_bruta == 0.0:
  for k, v in resumen.items():
    if k and str(k).strip().upper() in ["R. B.", "R.B.", "R.B"]:
      r_bruta = v
      break

# Fórmulas de Margen Bruto y Coste de Ventas
margen_bruto = r_bruta * ventas
coste_ventas = ventas - margen_bruto

otros_ingresos = get_val("Otros Ingresos")
ingresos_operativos = margen_bruto + otros_ingresos

gastos_personal = get_val("Gastos Personal")
alquileres = get_val("Alquileres")
reparaciones = get_val("Reparaciones")
seguros = get_val("Seguros")
suministros = get_val("Suministros")
otros_servicios = get_val("Otros Servicios")

total_gastos_operativos = (
    alquileres + reparaciones + seguros + suministros + otros_servicios
)
amortizaciones = get_val("Amortizaciones")

# Cálculo de Gastos de Estructura
gastos_estructura = total_gastos_operativos + gastos_personal + amortizaciones

baii = (
    ingresos_operativos
    - gastos_personal
    - total_gastos_operativos
    - amortizaciones
)

gastos_financieros = get_val("Gastos Financieros")
ingresos_financieros = get_val("Ingresos Financieros")
rdo_financiero = ingresos_financieros - gastos_financieros

resultados_extraordinarios = get_val("Resultados Extraordinarios")
bai = baii + rdo_financiero + resultados_extraordinarios

# Columna del mes/es para la tabla
col_nombre_periodo = f"{nombre_meses_str} {ano}"

data_out = [
    ("Ventas", ventas),
    ("Coste Ventas", coste_ventas),
    ("MARGEN BRUTO", margen_bruto),
    ("R. B.", r_bruta),
    ("Otros Ingresos", otros_ingresos),
    ("Ingresos Operativos", ingresos_operativos),
    ("Gastos Personal", gastos_personal),
    ("Alquileres", alquileres),
    ("Reparaciones", reparaciones),
    ("Seguros", seguros),
    ("Suministros", suministros),
    ("Otros Servicios", otros_servicios),
    ("TOTAL GASTOS OPERATIVOS", total_gastos_operativos),
    ("Amortizaciones", amortizaciones),
    ("GASTOS ESTRUCTURA", gastos_estructura),
    ("B.A.I.I.", baii),
    ("Gastos Financieros", gastos_financieros),
    ("Ingresos Financieros", ingresos_financieros),
    ("RDO. FINANCIERO", rdo_financiero),
    ("Resultados Extraordinarios", resultados_extraordinarios),
    ("B.A.I.", bai),
]

df_resultado = pd.DataFrame(data_out, columns=["Resultados", col_nombre_periodo])

# Creamos el DataFrame formateado para visualización
df_display = df_resultado.copy()


def formatear_valor(row):
  concepto = row["Resultados"]
  valor = row[col_nombre_periodo]
  if concepto == "R. B.":
    return (
        f"{valor * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    )
  else:
    return (
        f"{valor:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
    )


df_display[col_nombre_periodo] = df_display.apply(formatear_valor, axis=1)

# Estilizado avanzado con Pandas Styler: Concepto a la izquierda, Valores forzados rígidamente a la derecha
campos_destacados = [
    "MARGEN BRUTO",
    "R. B.",
    "Ingresos Operativos",
    "GASTOS ESTRUCTURA",
    "B.A.I.I.",
    "RDO. FINANCIERO",
    "B.A.I.",
]

df_styled = (
    df_display.style.set_properties(
        subset=["Resultados"],
        **{"text-align": "left", "padding-left": "10px"},
    )
    .set_properties(
        subset=[col_nombre_periodo],
        **{"text-align": "right", "padding-right": "20px"},
    )
    .apply(
        lambda row: [
            (
                "font-weight: bold; background-color: #eef2f7; color: #1f2937;"
                if row["Resultados"] in campos_destacados
                else "text-align: left;"
            )
        ]
        * 1
        + [
            (
                "font-weight: bold; background-color: #eef2f7; color: #1f2937;"
                if row["Resultados"] in campos_destacados
                else "text-align: right;"
            )
        ]
        * 1,
        axis=1,
    )
)

# Mostrar tabla estilizada en pantalla
st.dataframe(df_styled, use_container_width=True, hide_index=True)


# Botón de descarga directa en Excel (valores puros)
def to_excel(df_to_save):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_to_save.to_excel(writer, index=False, sheet_name="Informe")
  return output.getvalue()


excel_data = to_excel(df_resultado)
nombre_salida = (
    f"Informe_{'_'.join(tiendas)}_{nombre_meses_str.replace(', ', '_')}_{ano}.xlsx"
)

st.download_button(
    label="📥 Descargar Informe en Excel",
    data=excel_data,
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
