import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

# CSS ultra-optimizado para autoajustar y forzar alineaciones perfectas
st.markdown(
    """
    <style>
    /* Ocultar menú de Streamlit, footer y enlace a GitHub */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .viewerBadge_container {display: none !important;}
    a[href*="github.com"] {display: none !important;}
    
    /* Expandir la ventana al máximo y eliminar padding */
    .block-container {
        padding-top: 0.2rem !important;
        padding-bottom: 0.2rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    h1 {
        font-size: 1.2rem !important;
        margin-bottom: 0.1rem !important;
    }
    h3 {
        font-size: 0.95rem !important;
        margin-bottom: 0.1rem !important;
    }
    /* Estilo de tabla hipercompacta estilo Excel */
    table {
        width: 100% !important;
        font-size: 11px !important;
        border-collapse: collapse !important;
        table-layout: auto !important;
    }
    th, td {
        padding: 2px 6px !important;
        white-space: nowrap !important;
    }
    /* Columna de conceptos compacta y fija */
    th:first-child, td:first-child {
        width: 13% !important;
        text-align: left !important;
    }
    /* Columnas de datos con alineación absoluta a la derecha */
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

# Selector del Modo de Análisis
modo_analisis = st.sidebar.radio(
    "Tipo de Análisis",
    [
        "Evolución Mensual / Tienda",
        "Comparativa Multi-Tienda (Totales)",
        "Comparativa Interanual (Año vs Año Anterior)",
    ],
)

# Restringir años exclusivamente de 2024 a 2026
anos_disponibles = [2024, 2025, 2026]
if "Año" in df.columns:
  anos_excel = sorted(df["Año"].dropna().unique())
  anos_disponibles = [a for a in anos_disponibles if a in anos_excel]
if not anos_disponibles:
  anos_disponibles = [2026]

ano = st.sidebar.selectbox("Año principal", anos_disponibles)

# Filtrar departamentos exclusivamente del año seleccionado
df_ano = df[df["Año"] == ano] if "Año" in df.columns else df
departamentos_disponibles = (
    sorted(df_ano["Departamento"].dropna().unique())
    if "Departamento" in df_ano.columns
    else []
)

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

# Configuración específica según el modo de análisis
opcion_multitienda = "Total"
if modo_analisis == "Comparativa Multi-Tienda (Totales)":
  tiendas = st.sidebar.multiselect(
      "Selecciona tiendas a comparar",
      departamentos_disponibles,
      default=departamentos_disponibles[:2]
      if len(departamentos_disponibles) >= 2
      else departamentos_disponibles,
  )
  opcion_multitienda = st.sidebar.radio(
      "Columna final / Vista", ["Total", "Diferencias (Tienda 2 - Tienda 1)"]
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
  st.warning(
      "Por favor, selecciona al menos una tienda y un mes en la barra lateral."
  )
  st.stop()

nombre_meses_str = (
    ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} meses"
)
if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":
  st.subheader(
      f"Comparativa Interanual: {nombre_meses_str} ({ano} vs {ano - 1})"
  )
elif modo_analisis == "Comparativa Multi-Tienda (Totales)":
  st.subheader(f"Comparativa Multi-Tienda ({nombre_meses_str} {ano})")
else:
  st.subheader(f"Informe ({nombre_meses_str} {ano})")


# Función para calcular los resultados con agregación ponderada correcta de R.B.
def calcular_resultados(df_filtered):
  if df_filtered.empty:
    return {c: 0.0 for c in conceptos}

  resumen = df_filtered.groupby("Resultados")["Importe D"].sum().to_dict()

  def get_v(cat):
    return resumen.get(cat, 0.0)

  ventas = get_v("Ventas")

  total_mb = 0.0
  total_ventas_calc = 0.0
  for (ano_v, mes_v, dep_v), group in df_filtered.groupby(
      ["Año", "Mes", "Departamento"]
  ):
    v_row = group[group["Resultados"] == "Ventas"]["Importe D"].sum()
    rb_rows = group[
        group["Resultados"].str.strip().str.upper().isin(["R. B.", "R.B.", "R.B"])
    ]["Importe D"]
    rb_val = rb_rows.iloc[0] if not rb_rows.empty else 0.0
    total_ventas_calc += v_row
    total_mb += v_row * rb_val

  margen_bruto = total_mb
  r_bruta = (
      (margen_bruto / total_ventas_calc) if total_ventas_calc != 0 else 0.0
  )
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

datos_fuente = {}

if modo_analisis == "Comparativa Multi-Tienda (Totales)":
  if (
      opcion_multitienda == "Diferencias (Tienda 2 - Tienda 1)"
      and len(tiendas) >= 2
  ):
    columnas_eje = [tiendas[0], tiendas[1], "Var. €", "Var. %"]
  else:
    columnas_eje = tiendas.copy()
    if len(tiendas) > 1:
      columnas_eje.append("Total")

  for tienda in tiendas:
    mask = (
        (df["Año"] == ano)
        & (df["Mes"].isin(meses_sel))
        & (df["Departamento"] == tienda)
    )
    datos_fuente[tienda] = calcular_resultados(df[mask])

  if len(tiendas) > 1 and opcion_multitienda == "Total":
    mask_total = (
        (df["Año"] == ano)
        & (df["Mes"].isin(meses_sel))
        & (df["Departamento"].isin(tiendas))
    )
    datos_fuente["Total"] = calcular_resultados(df[mask_total])

elif modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":
  ano_anterior = ano - 1
  columnas_eje = [
      f"Total {ano}",
      f"Total {ano_anterior}",
      "Var. €",
      "Var. %",
  ]

  mask_ant = (
      (df["Año"] == ano_anterior)
      & (df["Mes"].isin(meses_sel))
      & (df["Departamento"].isin(tiendas))
  )
  mask_act = (
      (df["Año"] == ano)
      & (df["Mes"].isin(meses_sel))
      & (df["Departamento"].isin(tiendas))
  )

  datos_fuente["Ant"] = calcular_resultados(df[mask_ant])
  datos_fuente["Act"] = calcular_resultados(df[mask_act])

else:
  if len(tiendas) > 1:
    modo_comparativa = "tiendas"
    columnas_eje = tiendas.copy()
    columnas_eje.append("Total")

    for tienda in tiendas:
      mask = (
          (df["Año"] == ano)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"] == tienda)
      )
      datos_fuente[tienda] = calcular_resultados(df[mask])

    mask_total = (
        (df["Año"] == ano)
        & (df["Mes"].isin(meses_sel))
        & (df["Departamento"].isin(tiendas))
    )
    datos_fuente["Total"] = calcular_resultados(df[mask_total])
  else:
    modo_comparativa = "meses"
    columnas_eje = meses_sel.copy()
    if len(meses_sel) > 1:
      columnas_eje.append("Total")

    tienda_unica = tiendas[0]
    for mes in meses_sel:
      mask = (
          (df["Año"] == ano)
          & (df["Mes"] == mes)
          & (df["Departamento"] == tienda_unica)
      )
      datos_fuente[mes] = calcular_resultados(df[mask])

    if len(meses_sel) > 1:
      mask_total_meses = (
          (df["Año"] == ano)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"] == tienda_unica)
      )
      datos_fuente["Total"] = calcular_resultados(df[mask_total_meses])

columnas_tabla = ["Resultados"] + columnas_eje

filas_tabla_display = []
filas_valores_numericos = []

for concepto in conceptos:
  fila_disp = [concepto]
  fila_num = [concepto]

  if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":
    val_ant = datos_fuente["Ant"].get(concepto, 0.0)
    val_act = datos_fuente["Act"].get(concepto, 0.0)

    if concepto == "R. B.":
      var_diff = val_act - val_ant
      fila_num.extend([val_act, val_ant, var_diff, 0.0])
      fila_disp.append(
          f"{val_act * 100:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{val_ant * 100:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{var_diff * 100:,.2f} pp"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append("-")
    else:
      var_eur = val_act - val_ant
      var_pct = (var_eur / abs(val_ant) * 100) if val_ant != 0 else 0.0
      fila_num.extend([val_act, val_ant, var_eur, var_pct])

      fila_disp.append(
          f"{val_act:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{val_ant:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{var_eur:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{var_pct:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )

  elif (
      modo_analisis == "Comparativa Multi-Tienda (Totales)"
      and opcion_multitienda == "Diferencias (Tienda 2 - Tienda 1)"
      and len(tiendas) >= 2
  ):
    val_t1 = datos_fuente[tiendas[0]].get(concepto, 0.0)
    val_t2 = datos_fuente[tiendas[1]].get(concepto, 0.0)

    if concepto == "R. B.":
      var_diff = val_t2 - val_t1
      fila_num.extend([val_t1, val_t2, var_diff, 0.0])
      fila_disp.append(
          f"{val_t1 * 100:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{val_t2 * 100:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{var_diff * 100:,.2f} pp"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append("-")
    else:
      var_eur = val_t2 - val_t1
      var_pct = (var_eur / abs(val_t1) * 100) if val_t1 != 0 else 0.0
      fila_num.extend([val_t1, val_t2, var_eur, var_pct])

      fila_disp.append(
          f"{val_t1:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{val_t2:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{var_eur:,.2f} €"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )
      fila_disp.append(
          f"{var_pct:,.2f}%"
          .replace(",", "X")
          .replace(".", ",")
          .replace("X", ".")
      )

  else:
    ejes_eval = (
        tiendas
        if modo_analisis == "Comparativa Multi-Tienda (Totales)"
        else (
            tiendas
            if (len(tiendas) > 1 and modo_comparativa == "tiendas")
            else meses_sel
        )
    )

    for item in ejes_eval:
      val = datos_fuente[item].get(concepto, 0.0)
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

    if len(columnas_eje) > len(ejes_eval):
      if concepto == "R. B.":
        if modo_analisis == "Comparativa Multi-Tienda (Totales)":
          tot_ventas = sum(datos_fuente[t].get("Ventas", 0.0) for t in tiendas)
          tot_margen = sum(
              datos_fuente[t].get("MARGEN BRUTO", 0.0) for t in tiendas
          )
        elif len(tiendas) > 1:
          tot_ventas = sum(datos_fuente[t].get("Ventas", 0.0) for t in tiendas)
          tot_margen = sum(
              datos_fuente[t].get("MARGEN BRUTO", 0.0) for t in tiendas
          )
        else:
          tot_ventas = sum(datos_fuente[m].get("Ventas", 0.0) for m in meses_sel)
          tot_margen = sum(
              datos_fuente[m].get("MARGEN BRUTO", 0.0) for m in meses_sel
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
        if modo_analisis == "Comparativa Multi-Tienda (Totales)":
          val_total = sum(datos_fuente[t].get(concepto, 0.0) for t in tiendas)
        elif len(tiendas) > 1:
          val_total = sum(datos_fuente[t].get(concepto, 0.0) for t in tiendas)
        else:
          val_total = sum(datos_fuente[m].get(concepto, 0.0) for m in meses_sel)
        fila_num.append(val_total)
        fila_disp.append(
            f"{val_total:,.2f} €"
            .replace(",", "X").replace(".", ",").replace("X", ".")
        )

  filas_tabla_display.append(fila_disp)
  filas_valores_numericos.append(fila_num)

df_resultado_display = pd.DataFrame(
    filas_tabla_display, columns=columnas_tabla
)
df_valores_numericos = pd.DataFrame(
    filas_valores_numericos, columns=columnas_tabla
)

campos_destacados = [
    "MARGEN BRUTO",
    "R. B.",
    "Ingresos Operativos",
    "GASTOS ESTRUCTURA",
    "B.A.I.I.",
    "RDO. FINANCIERO",
    "B.A.I.",
]

# Incluimos R. B. como concepto tipo ingreso (positivo = verde, negativo = rojo)
conceptos_ingresos = [
    "Ventas",
    "MARGEN BRUTO",
    "R. B.",
    "Otros Ingresos",
    "Ingresos Operativos",
    "Ingresos Financieros",
    "B.A.I.I.",
    "RDO. FINANCIERO",
    "B.A.I.",
]


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
      is_columna_total = col == "Total" or col == f"Total {ano}"

      cell_style = "text-align: right !important; padding-right: 6px;"

      if is_columna_total:
        cell_style += " background-color: #d1fae5;"
      elif is_destacado:
        cell_style += " background-color: #eef2f7;"

      if is_destacado:
        cell_style += " font-weight: bold;"

      if col in ["Var. €", "Var. %"] and isinstance(num_val, (int, float)):
        if num_val != 0:
          es_ingreso = concepto in conceptos_ingresos
          if es_ingreso:
            color_var = "#16a34a" if num_val > 0 else "#dc2626"
          else:
            color_var = "#dc2626" if num_val > 0 else "#16a34a"
          cell_style += f" color: {color_var}; font-weight: bold;"
      else:
        if is_negativo:
          cell_style += " color: #dc2626; font-weight: bold;"
        elif is_destacado:
          cell_style += " color: #1f2937;"

      row_styles.append(cell_style)
    styles.append(row_styles)

  return pd.DataFrame(styles, index=s.index, columns=s.columns)


df_styled = df_resultado_display.style.apply(aplicar_estilos_styler, axis=None)

# Mostrar la tabla estática con autoajuste fluido
st.table(df_styled)


def to_excel(df_to_save):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_to_save.to_excel(writer, index=False, sheet_name="Informe")
  return output.getvalue()


excel_data = to_excel(df_valores_numericos)
nombre_salida = f"Informe_Resultados_{ano}.xlsx"

st.download_button(
    label="📥 Descargar Informe en Excel",
    data=excel_data,
    file_name=nombre_salida,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
