import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Control de Resultados - Herederos", layout="wide")

# CSS optimizado para adaptabilidad perfecta de columnas
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
    
    /* Estilos limpios para tablas nativas ajustadas */
    div[data-testid="stDataFrame"], div.stDataFrame {
        width: 100% !important;
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

# Selector del Módulo Principal
modulo_principal = st.sidebar.radio(
    "Módulo de Análisis",
    [
        "Cuenta de Resultados Completa",
        "Análisis Específico de R.B. (Margen Bruto)",
        "Informe KPI (% sobre Ventas)",
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

  def calcular_rb_puro(df_f):
    if df_f.empty:
      return 0.0
    total_mb = 0.0
    total_ventas = 0.0
    for (a_v, m_v, d_v), group in df_f.groupby(["Año", "Mes", "Departamento"]):
      v_row = group[group["Resultados"] == "Ventas"]["Importe D"].sum()
      rb_rows = group[
          group["Resultados"]
          .str.strip()
          .str.upper()
          .isin(["R. B.", "R.B.", "R.B"])
      ]["Importe D"]
      rb_val = rb_rows.iloc[0] if not rb_rows.empty else 0.0
      total_ventas += v_row
      total_mb += v_row * rb_val
    if total_ventas == 0:
      resumen = df_f.groupby("Resultados")["Importe D"].sum().to_dict()
      return resumen.get("R. B.", 0.0)
    return total_mb / total_ventas

  if tipo_analisis_rb == "Evolución Mensual por Tienda":
    tiendas_rb = st.sidebar.multiselect(
        "Selecciona tiendas",
        departamentos_disponibles,
        default=departamentos_disponibles,
    )
    if not tiendas_rb or not meses_sel:
      st.warning("Selecciona al menos una tienda y un mes.")
      st.stop()

    st.subheader(
        f"Análisis R.B. - Evolución Mensual por Tienda ({ano})"
    )

    columnas_tabla = ["Resultados"] + meses_sel
    if len(meses_sel) > 1:
      columnas_tabla.append("Promedio Acumulado")

    filas_display = []
    filas_nums = []

    for tienda in tiendas_rb:
      fila_d = [tienda]
      fila_n = [tienda]
      for mes in meses_sel:
        mask = (
            (df["Año"] == ano)
            & (df["Mes"] == mes)
            & (df["Departamento"] == tienda)
        )
        val_rb = calcular_rb_puro(df[mask])
        fila_n.append(val_rb)
        fila_d.append(
            f"{val_rb * 100:,.2f}%"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

      if len(meses_sel) > 1:
        mask_acum = (
            (df["Año"] == ano)
            & (df["Mes"].isin(meses_sel))
            & (df["Departamento"] == tienda)
        )
        val_acum = calcular_rb_puro(df[mask_acum])
        fila_n.append(val_acum)
        fila_d.append(
            f"{val_acum * 100:,.2f}%"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

      filas_display.append(fila_d)
      filas_nums.append(fila_n)

    if len(tiendas_rb) > 1:
      fila_d_tot = ["TOTAL GRUPO"]
      fila_n_tot = ["TOTAL GRUPO"]
      for mes in meses_sel:
        mask_m = (
            (df["Año"] == ano)
            & (df["Mes"] == mes)
            & (df["Departamento"].isin(tiendas_rb))
        )
        val_m = calcular_rb_puro(df[mask_m])
        fila_n_tot.append(val_m)
        fila_d_tot.append(
            f"{val_m * 100:,.2f}%"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
      if len(meses_sel) > 1:
        mask_tot_acum = (
            (df["Año"] == ano)
            & (df["Mes"].isin(meses_sel))
            & (df["Departamento"].isin(tiendas_rb))
        )
        val_tot_ac = calcular_rb_puro(df[mask_tot_acum])
        fila_n_tot.append(val_tot_ac)
        fila_d_tot.append(
            f"{val_tot_ac * 100:,.2f}%"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
      filas_display.append(fila_d_tot)
      filas_nums.append(fila_n_tot)

    df_res_d = pd.DataFrame(filas_display, columns=columnas_tabla)
    df_res_n = pd.DataFrame(filas_nums, columns=columnas_tabla)

    def estilizar_rb(s):
      styles = []
      for i, row in df_res_d.iterrows():
        is_total = row["Resultados"] == "TOTAL GRUPO"
        row_styles = [
            "text-align: left !important; padding-left: 6px;"
            + ("font-weight: bold; background-color: #d1fae5;" if is_total else "")
        ]
        for col in columnas_tabla[1:]:
          c_style = "text-align: right !important; padding-right: 8px;"
          if is_total or col == "Promedio Acumulado":
            c_style += " background-color: #d1fae5; font-weight: bold;"
          row_styles.append(c_style)
        styles.append(row_styles)
      return pd.DataFrame(styles, index=s.index, columns=s.columns)

    st.dataframe(
        df_res_d.style.apply(estilizar_rb, axis=None),
        use_container_width=True,
        hide_index=True,
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_res_n.to_excel(writer, index=False, sheet_name="Analisis_RB_Mensual")
    st.download_button(
        label="📥 Descargar Análisis R.B. en Excel",
        data=output.getvalue(),
        file_name=f"Analisis_RB_Mensual_{ano}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  elif tipo_analisis_rb == "Vista Acumulada por Tienda":
    tiendas_rb = st.sidebar.multiselect(
        "Selecciona tiendas",
        departamentos_disponibles,
        default=departamentos_disponibles,
    )
    if not tiendas_rb or not meses_sel:
      st.warning("Selecciona al menos una tienda y un mes.")
      st.stop()

    nombre_m_str = (
        ", ".join(meses_sel)
        if len(meses_sel) <= 3
        else f"{len(meses_sel)} meses acumulados"
    )
    st.subheader(f"Análisis R.B. - Vista Acumulada ({nombre_m_str} {ano})")

    filas_acum_d = []
    filas_acum_n = []

    for tienda in tiendas_rb:
      mask = (
          (df["Año"] == ano)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"] == tienda)
      )
      val_rb = calcular_rb_puro(df[mask])
      filas_acum_d.append([tienda, f"{val_rb * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")])
      filas_acum_n.append([tienda, val_rb])

    if len(tiendas_rb) > 1:
      mask_tot = (
          (df["Año"] == ano)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"].isin(tiendas_rb))
      )
      val_tot = calcular_rb_puro(df[mask_tot])
      filas_acum_d.append(["TOTAL GRUPO", f"{val_tot * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")])
      filas_acum_n.append(["TOTAL GRUPO", val_tot])

    df_acum_d = pd.DataFrame(filas_acum_d, columns=["Resultados", f"Acumulado {nombre_m_str}"])
    df_acum_n = pd.DataFrame(filas_acum_n, columns=["Resultados", f"Acumulado {nombre_m_str}"])

    def estilizar_acum(s):
      styles = []
      for i, row in df_acum_d.iterrows():
        is_tot = row["Resultados"] == "TOTAL GRUPO"
        row_styles = [
            "text-align: left !important; padding-left: 6px;"
            + ("font-weight: bold; background-color: #d1fae5;" if is_tot else "")
        ]
        c_style = "text-align: right !important; padding-right: 8px;"
        if is_tot:
          c_style += " background-color: #d1fae5; font-weight: bold;"
        row_styles.append(c_style)
        styles.append(row_styles)
      return pd.DataFrame(styles, index=s.index, columns=s.columns)

    st.dataframe(
        df_acum_d.style.apply(estilizar_acum, axis=None),
        use_container_width=True,
        hide_index=True,
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_acum_n.to_excel(writer, index=False, sheet_name="Analisis_RB_Acumulado")
    st.download_button(
        label="📥 Descargar Acumulado R.B. en Excel",
        data=output.getvalue(),
        file_name=f"Analisis_RB_Acumulado_{ano}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  else:  # Comparativa Interanual R.B.
    tiendas_rb = st.sidebar.multiselect(
        "Selecciona tiendas",
        departamentos_disponibles,
        default=departamentos_disponibles,
    )
    if not tiendas_rb or not meses_sel:
      st.warning("Selecciona al menos una tienda y un mes.")
      st.stop()

    ano_ant = ano - 1
    nombre_m_str = (
        ", ".join(meses_sel)
        if len(meses_sel) <= 3
        else f"{len(meses_sel)} meses"
    )
    st.subheader(
        f"Comparativa Interanual R.B. ({nombre_m_str}): {ano} vs {ano_ant}"
    )

    col_a = f"R.B. {ano}"
    col_b = f"R.B. {ano_ant}"
    columnas_interanual_rb = ["Resultados", col_a, col_b, "Var. pp"]

    filas_inter_d = []
    filas_inter_n = []

    for tienda in tiendas_rb:
      mask_act = (
          (df["Año"] == ano)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"] == tienda)
      )
      mask_ant = (
          (df["Año"] == ano_ant)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"] == tienda)
      )

      val_act = calcular_rb_puro(df[mask_act])
      val_ant = calcular_rb_puro(df[mask_ant])
      var_pp = val_act - val_ant

      filas_inter_n.append([tienda, val_act, val_ant, var_pp])
      filas_inter_d.append([
          tienda,
          f"{val_act * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."),
          f"{val_ant * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."),
          f"{var_pp * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", "."),
      ])

    if len(tiendas_rb) > 1:
      mask_tot_act = (
          (df["Año"] == ano)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"].isin(tiendas_rb))
      )
      mask_tot_ant = (
          (df["Año"] == ano_ant)
          & (df["Mes"].isin(meses_sel))
          & (df["Departamento"].isin(tiendas_rb))
      )
      val_tot_act = calcular_rb_puro(df[mask_tot_act])
      val_tot_ant = calcular_rb_puro(df[mask_tot_ant])
      var_tot_pp = val_tot_act - val_tot_ant

      filas_inter_n.append(
          ["TOTAL GRUPO", val_tot_act, val_tot_ant, var_tot_pp]
      )
      filas_inter_d.append([
          "TOTAL GRUPO",
          f"{val_tot_act * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."),
          f"{val_tot_ant * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."),
          f"{var_tot_pp * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", "."),
      ])

    df_inter_d = pd.DataFrame(filas_inter_d, columns=columnas_interanual_rb)
    df_inter_n = pd.DataFrame(filas_inter_n, columns=columnas_interanual_rb)

    def estilizar_inter(s):
      styles = []
      for i, row in df_inter_d.iterrows():
        is_tot = row["Resultados"] == "TOTAL GRUPO"
        row_styles = [
            "text-align: left !important; padding-left: 6px;"
            + ("font-weight: bold; background-color: #d1fae5;" if is_tot else "")
        ]
        for col_idx, col in enumerate(columnas_interanual_rb[1:], start=1):
          c_style = "text-align: right !important; padding-right: 8px;"
          if is_tot:
            c_style += " background-color: #d1fae5; font-weight: bold;"
          if col == "Var. pp":
            val_num = filas_inter_n[i][col_idx]
            if val_num != 0:
              color_v = "#16a34a" if val_num > 0 else "#dc2626"
              c_style += f" color: {color_v}; font-weight: bold;"
          row_styles.append(c_style)
        styles.append(row_styles)
      return pd.DataFrame(styles, index=s.index, columns=s.columns)

    st.dataframe(
        df_inter_d.style.apply(estilizar_inter, axis=None),
        use_container_width=True,
        hide_index=True,
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_inter_n.to_excel(writer, index=False, sheet_name="Interanual_RB")
    st.download_button(
        label="📥 Descargar Comparativa R.B. en Excel",
        data=output.getvalue(),
        file_name=f"Comparativa_Interanual_RB_{ano}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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

  df_ano = df[df["Año"] == ano] if "Año" in df.columns else df
  departamentos_disponibles = (
      sorted(df_ano["Departamento"].dropna().unique())
      if "Departamento" in df_ano.columns
      else []
  )

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
    st.warning("Selecciona al menos una tienda y un mes.")
    st.stop()

  nombre_meses_str = (
      ", ".join(meses_sel) if len(meses_sel) <= 3 else f"{len(meses_sel)} meses"
  )
  if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":
    st.subheader(
        f"Informe KPI (% sobre Ventas) - Interanual: {nombre_meses_str} ({ano} vs {ano - 1})"
    )
  elif modo_analisis == "Comparativa Multi-Tienda (Totales)":
    st.subheader(
        f"Informe KPI (% sobre Ventas) - Multi-Tienda ({nombre_meses_str} {ano})"
    )
  else:
    st.subheader(
        f"Informe KPI (% sobre Ventas) - ({nombre_meses_str} {ano})"
    )

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
          group["Resultados"]
          .str.strip()
          .str.upper()
          .isin(["R. B.", "R.B.", "R.B"])
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
      columnas_eje = [tiendas[0], tiendas[1], "Var. pp"]
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
    columnas_eje = [f"Total {ano}", f"Total {ano_anterior}", "Var. pp"]
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

  conceptos_gastos = [
      "Coste Ventas", "Gastos Personal", "Alquileres", "Reparaciones",
      "Seguros", "Suministros", "Otros Servicios", "TOTAL GASTOS OPERATIVOS",
      "Amortizaciones", "GASTOS ESTRUCTURA", "Gastos Financieros"
  ]

  for concepto in conceptos:
    fila_disp = [concepto]
    fila_num = [concepto]

    if modo_analisis == "Comparativa Interanual (Año vs Año Anterior)":
      res_ant = datos_fuente["Ant"]
      res_act = datos_fuente["Act"]
      v_ventas_act = res_act.get("Ventas", 0.0)
      v_ventas_ant = res_ant.get("Ventas", 0.0)

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
      fila_disp.append(f"{kpi_act * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
      fila_disp.append(f"{kpi_ant * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
      fila_disp.append(f"{var_pp * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", "."))

    elif (
        modo_analisis == "Comparativa Multi-Tienda (Totales)"
        and opcion_multitienda == "Diferencias (Tienda 2 - Tienda 1)"
        and len(tiendas) >= 2
    ):
      res_t1 = datos_fuente[tiendas[0]]
      res_t2 = datos_fuente[tiendas[1]]
      v_v1 = res_t1.get("Ventas", 0.0)
      v_v2 = res_t2.get("Ventas", 0.0)

      val_t1_abs = res_t1.get(concepto, 0.0)
      val_t2_abs = res_t2.get(concepto, 0.0)

      if concepto == "R. B.":
        kpi_t1 = val_t1_abs
        kpi_t2 = val_t2_abs
      else:
        kpi_t1 = (val_t1_abs / v_v1) if v_v1 != 0 else 0.0
        kpi_t2 = (val_t2_abs / v_v2) if v_v2 != 0 else 0.0

      var_pp = kpi_t2 - kpi_t1
      fila_num.extend([kpi_t1, kpi_t2, var_pp])
      fila_disp.append(f"{kpi_t1 * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
      fila_disp.append(f"{kpi_t2 * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
      fila_disp.append(f"{var_pp * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", "."))

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
        res_item = datos_fuente[item]
        v_ventas_item = res_item.get("Ventas", 0.0)
        val_abs = res_item.get(concepto, 0.0)

        if concepto == "R. B.":
          kpi_val = val_abs
        else:
          kpi_val = (val_abs / v_ventas_item) if v_ventas_item != 0 else 0.0

        fila_num.append(kpi_val)
        fila_disp.append(f"{kpi_val * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))

      if len(columnas_eje) > len(ejes_eval):
        if modo_analisis == "Comparativa Multi-Tienda (Totales)":
          tot_ventas = sum(datos_fuente[t].get("Ventas", 0.0) for t in tiendas)
          tot_concepto = sum(datos_fuente[t].get(concepto, 0.0) for t in tiendas)
        elif len(tiendas) > 1:
          tot_ventas = sum(datos_fuente[t].get("Ventas", 0.0) for t in tiendas)
          tot_concepto = sum(datos_fuente[t].get(concepto, 0.0) for t in tiendas)
        else:
          tot_ventas = sum(datos_fuente[m].get("Ventas", 0.0) for m in meses_sel)
          tot_concepto = sum(datos_fuente[m].get(concepto, 0.0) for m in meses_sel)

        if concepto == "R. B.":
          tot_mb_val = 0.0
          for itm in (tiendas if modo_analisis == "Comparativa Multi-Tienda (Totales)" else (tiendas if len(tiendas)>1 else meses_sel)):
            v_v = datos_fuente[itm].get("Ventas", 0.0)
            rb_v = datos_fuente[itm].get("R. B.", 0.0)
            tot_mb_val += v_v * rb_v
          kpi_total = (tot_mb_val / tot_ventas) if tot_ventas != 0 else 0.0
        else:
          kpi_total = (tot_concepto / tot_ventas) if tot_ventas != 0 else 0.0

        fila_num.append(kpi_total)
        fila_disp.append(f"{kpi_total * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))

    filas_tabla_display.append(fila_disp)
    filas_valores_numericos.append(fila_num)

  df_kpi_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
  df_kpi_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)

  campos_destacados = [
      "MARGEN BRUTO",
      "R. B.",
      "Ingresos Operativos",
      "TOTAL GASTOS OPERATIVOS",
      "GASTOS ESTRUCTURA",
      "B.A.I.I.",
      "RDO. FINANCIERO",
      "B.A.I.",
  ]

  def aplicar_estilos_kpi(s):
    styles = []
    es_multi_tienda = (modo_analisis == "Comparativa Multi-Tienda (Totales)" and len(tiendas) > 1 and opcion_multitienda == "Total") or \
                      (modo_analisis == "Evolución Mensual / Tienda" and len(tiendas) > 1)

    for i, row in df_kpi_display.iterrows():
      concepto = row["Resultados"]
      is_destacado = concepto in campos_destacados
      row_styles = [
          "text-align: left !important; padding-left: 6px;"
          + ("font-weight: bold; background-color: #eef2f7;" if is_destacado else "")
      ]

      tiendas_valores = []
      if es_multi_tienda:
        for col_idx, col in enumerate(columnas_tabla[1:], start=1):
          if col != "Total":
            val_n = df_kpi_numericos.loc[i, col]
            if isinstance(val_n, (int, float)):
              tiendas_valores.append(val_n)

      max_val = max(tiendas_valores) if tiendas_valores else None
      min_val = min(tiendas_valores) if tiendas_valores else None

      for col_idx, col in enumerate(columnas_tabla[1:], start=1):
        num_val = df_kpi_numericos.loc[i, col]
        is_negativo = isinstance(num_val, (int, float)) and num_val < 0
        is_columna_total = col == "Total" or col == f"Total {ano}"
        cell_style = "text-align: right !important; padding-right: 8px;"

        if is_columna_total:
          cell_style += " background-color: #d1fae5;"
        elif is_destacado:
          cell_style += " background-color: #eef2f7;"
        if is_destacado:
          cell_style += " font-weight: bold;"

        if es_multi_tienda and col != "Total" and max_val is not None and min_val is not None and max_val != min_val:
          if isinstance(num_val, (int, float)):
            if concepto in ["RDO. FINANCIERO", "Resultados Extraordinarios"]:
              if num_val == min_val:
                cell_style += " color: #16a34a; background-color: #d1fae5; font-weight: bold;"
              elif num_val == max_val:
                cell_style += " color: #dc2626; background-color: #fee2e2; font-weight: bold;"
            elif concepto in conceptos_gastos:
              if num_val == min_val:
                cell_style += " color: #16a34a; background-color: #d1fae5; font-weight: bold;"
              elif num_val == max_val:
                cell_style += " color: #dc2626; background-color: #fee2e2; font-weight: bold;"
            else:
              if num_val == max_val:
                cell_style += " color: #16a34a; background-color: #d1fae5; font-weight: bold;"
              elif num_val == min_val:
                cell_style += " color: #dc2626; background-color: #fee2e2; font-weight: bold;"

        if col == "Var. pp" and isinstance(num_val, (int, float)):
          if num_val != 0:
            color_v = "#16a34a" if num_val > 0 else "#dc2626"
            cell_style += f" color: {color_v}; font-weight: bold;"
        else:
          if is_negativo and not (es_multi_tienda and col != "Total"):
            cell_style += " color: #dc2626; font-weight: bold;"

        row_styles.append(cell_style)
      styles.append(row_styles)
    return pd.DataFrame(styles, index=s.index, columns=s.columns)

  st.dataframe(
      df_kpi_display.style.apply(aplicar_estilos_kpi, axis=None),
      use_container_width=True,
      hide_index=True,
  )

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_kpi_numericos.to_excel(writer, index=False, sheet_name="Informe_KPI")
  st.download_button(
      label="📥 Descargar Informe KPI en Excel",
      data=output.getvalue(),
      file_name=f"Informe_KPI_Ventas_{ano}.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )


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

  df_ano = df[df["Año"] == ano] if "Año" in df.columns else df
  departamentos_disponibles = (
      sorted(df_ano["Departamento"].dropna().unique())
      if "Departamento" in df_ano.columns
      else []
  )

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
    st.warning("Selecciona al menos una tienda y un mes.")
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
          group["Resultados"]
          .str.strip()
          .str.upper()
          .isin(["R. B.", "R.B.", "R.B"])
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
    columnas_eje = [f"Total {ano}", f"Total {ano_anterior}", "Var. €", "Var. %"]
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
        fila_disp.append(f"{val_act * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{val_ant * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{var_diff * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append("-")
      else:
        var_eur = val_act - val_ant
        var_pct = (var_eur / abs(val_ant) * 100) if val_ant != 0 else 0.0
        fila_num.extend([val_act, val_ant, var_eur, var_pct])
        fila_disp.append(f"{val_act:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{val_ant:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{var_eur:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{var_pct:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))

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
        fila_disp.append(f"{val_t1 * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{val_t2 * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{var_diff * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append("-")
      else:
        var_eur = val_t2 - val_t1
        var_pct = (var_eur / abs(val_t1) * 100) if val_t1 != 0 else 0.0
        fila_num.extend([val_t1, val_t2, var_eur, var_pct])
        fila_disp.append(f"{val_t1:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{val_t2:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{var_eur:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        fila_disp.append(f"{var_pct:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))

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
          fila_disp.append(f"{val * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
          fila_disp.append(f"{val:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))

      if len(columnas_eje) > len(ejes_eval):
        if concepto == "R. B.":
          if modo_analisis == "Comparativa Multi-Tienda (Totales)":
            tot_ventas = sum(datos_fuente[t].get("Ventas", 0.0) for t in tiendas)
            tot_margen = sum(datos_fuente[t].get("MARGEN BRUTO", 0.0) for t in tiendas)
          elif len(tiendas) > 1:
            tot_ventas = sum(datos_fuente[t].get("Ventas", 0.0) for t in tiendas)
            tot_margen = sum(datos_fuente[t].get("MARGEN BRUTO", 0.0) for t in tiendas)
          else:
            tot_ventas = sum(datos_fuente[m].get("Ventas", 0.0) for m in meses_sel)
            tot_margen = sum(datos_fuente[m].get("MARGEN BRUTO", 0.0) for m in meses_sel)
          val_total = (tot_margen / tot_ventas) if tot_ventas != 0 else 0.0
          fila_num.append(val_total)
          fila_disp.append(f"{val_total * 100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
          if modo_analisis == "Comparativa Multi-Tienda (Totales)":
            val_total = sum(datos_fuente[t].get(concepto, 0.0) for t in tiendas)
          elif len(tiendas) > 1:
            val_total = sum(datos_fuente[t].get(concepto, 0.0) for t in tiendas)
          else:
            val_total = sum(datos_fuente[m].get(concepto, 0.0) for m in meses_sel)
          fila_num.append(val_total)
          fila_disp.append(f"{val_total:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))

    filas_tabla_display.append(fila_disp)
    filas_valores_numericos.append(fila_num)

  df_resultado_display = pd.DataFrame(filas_tabla_display, columns=columnas_tabla)
  df_valores_numericos = pd.DataFrame(filas_valores_numericos, columns=columnas_tabla)

  campos_destacados = [
      "MARGEN BRUTO",
      "R. B.",
      "Ingresos Operativos",
      "TOTAL GASTOS OPERATIVOS",
      "GASTOS ESTRUCTURA",
      "B.A.I.I.",
      "RDO. FINANCIERO",
      "B.A.I.",
  ]
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
        cell_style = "text-align: right !important; padding-right: 8px;"
        if is_columna_total:
          cell_style += " background-color: #d1fae5;"
        elif is_destacado:
          cell_style += " background-color: #eef2f7;"
        if is_destacado:
          cell_style += " font-weight: bold;"
        if col in ["Var. €", "Var. %"] and isinstance(num_val, (int, float)):
          if num_val != 0:
            es_ingreso = concepto in conceptos_ingresos
            color_var = "#16a34a" if (num_val > 0 if es_ingreso else num_val < 0) else "#dc2626"
            cell_style += f" color: {color_var}; font-weight: bold;"
        else:
          if is_negativo:
            cell_style += " color: #dc2626; font-weight: bold;"
          elif is_destacado:
            cell_style += " color: #1f2937;"
        row_styles.append(cell_style)
      styles.append(row_styles)
    return pd.DataFrame(styles, index=s.index, columns=s.columns)

  st.dataframe(
      df_resultado_display.style.apply(aplicar_estilos_styler, axis=None),
      use_container_width=True,
      hide_index=True,
  )

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_valores_numericos.to_excel(writer, index=False, sheet_name="Informe")
  st.download_button(
      label="📥 Descargar Informe en Excel",
      data=output.getvalue(),
      file_name=f"Informe_Resultados_{ano}.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )
