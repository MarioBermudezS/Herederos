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

anos_disponibles = (
    sorted(df["Año"].dropna().unique()) if "Año" in df.columns else [2026]
)
ano = st.sidebar.selectbox("Año", anos_disponibles)

meses_disponibles = (
    df["Mes"].dropna().unique().tolist() if "Mes" in df.columns else ["Enero"]
)
mes = st.sidebar.selectbox("Mes", meses_disponibles)

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

if not tiendas:
  st.warning("Por favor, selecciona al menos una tienda en la barra lateral.")
  st.stop()

st.subheader(f"Informe para: {', '.join(tiendas)} ({mes} {ano})")

# Filtrado de datos
mask = (df["Año"] == ano) & (df["Mes"] == mes) & (df["Departamento"].isin(tiendas))
df_filtered = df[mask]

if df_filtered.empty:
  st.warning("¡Aviso! No se han encontrado datos para esos criterios.")
  st.stop()

resumen = df_filtered.groupby("Resultados")["Importe D"].sum().to_dict()


def get_val(cat):
  return resumen.get(cat, 0.0)


ventas = get_val("Ventas")
coste_ventas = get_val("Consumo Ventas")
margen_bruto = ventas - coste_ventas
r_bruta = (margen_bruto / ventas) if ventas != 0 else 0.0

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

data_out = [
    ("Ventas", ventas),
    ("Coste Ventas", coste_ventas),
    ("MARGEN BRUTO", margen_bruto),
    ("R. Bruta", r_bruta),
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
    ("B.A.I.I.", baii),
    ("Gastos Financieros", gastos_financieros),
    ("Ingresos Financieros", ingresos_financieros),
    ("RDO. FINANCIERO", rdo_financiero),
    ("Resultados Extraordinarios", resultados_extraordinarios),
    ("B.A.I.", bai),
]

df_resultado = pd.DataFrame(data_out, columns=["Resultados", f"{mes} {ano}"])

# Mostrar tabla en pantalla
st.dataframe(df_resultado, use_container_width=True, hide_index=True)


# Botón de descarga directa en Excel
def to_excel(df_to_save):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_to_save.to_excel(writer, index=False, sheet_name="Informe")
  return output.getvalue()


excel_data = to_excel(df_resultado)
nombre_salida = f"Informe_{'_'.join(tiendas)}_{mes}_{ano}.xlsx"

st.download_button(
    label="📥 Descargar Informe en Excel",
    data=excel_data,
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
