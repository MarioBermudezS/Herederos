campos_destacados = [
    "MARGEN BRUTO",
    "R. B.",
    "Ingresos Operativos",
    "TOTAL GASTOS OPERATIVOS",
    "GASTOS ESTRUCTURA",
    "B.A.I.I.",
    "RDO. FINANCIERO",
    "Resultados Extraordinarios",
    "B.A.I.",
    "TOTAL GRUPO",
]

# JsCode avanzado para estilos, negritas, totales y colores condicionales en KPIs
cell_style_jscode = JsCode("""
function(params) {
    var rowNode = params.node;
    var colDef = params.colDef;
    var val = params.value;
    var field = colDef.field;
    var rowLabel = rowNode.data.Resultados || '';

    var isDestacado = [
        "MARGEN BRUTO", "R. B.", "Ingresos Operativos", 
        "TOTAL GASTOS OPERATIVOS", "GASTOS ESTRUCTURA", 
        "B.A.I.I.", "RDO. FINANCIERO", "Resultados Extraordinarios", "B.A.I.", "TOTAL GRUPO"
    ].includes(rowLabel);

    var isTotalCol = field === "Total" || field.startsWith("Total ");
    var isFirstCol = colDef.pinned === "left" || colDef.field === "Resultados";

    var style = {
        'textAlign': isFirstCol ? 'left' : 'right',
        'fontWeight': (isDestacado || isTotalCol) ? 'bold' : 'normal'
    };

    if (isTotalCol) {
        style['backgroundColor'] = '#d1fae5';
    } else if (isDestacado) {
        style['backgroundColor'] = '#eef2f7';
    }

    if (typeof val === 'string') {
        if (val.includes('-') && !val.includes('%') && !val.includes('pp')) {
            style['color'] = '#dc2626';
            style['fontWeight'] = 'bold';
        } else if (field === "Var. pp" || field === "Var. %" || field === "Var. €") {
            if (!val.includes('-') && val !== '-' && val !== '0,00%' && val !== '0,00 pp') {
                style['color'] = '#16a34a';
                style['fontWeight'] = 'bold';
            } else if (val.includes('-')) {
                style['color'] = '#dc2626';
                style['fontWeight'] = 'bold';
            }
        }
    }

    return style;
}
""")


def render_tabla_aggrid(df_display):
  gb = GridOptionsBuilder.from_dataframe(df_display)
  gb.configure_default_column(
      resizable=True,
      filterable=False,
      sortable=False,
      editable=False,
      suppressMenu=True,
      cellStyle=cell_style_jscode,
  )

  if len(df_display.columns) > 0:
    first_col = df_display.columns[0]
    gb.configure_column(
        first_col,
        pinned="left",
        width=230,
        minWidth=200,
    )

  for col in df_display.columns[1:]:
    gb.configure_column(col, width=125, minWidth=105)

  gb.configure_grid_options(
      suppressRowClickSelection=True,
  )
  gridOptions = gb.build()

  row_count = len(df_display)
  calculated_height = max(400, (row_count + 1) * 36 + 45)

  AgGrid(
      df_display,
      gridOptions=gridOptions,
      height=calculated_height,
      update_mode=GridUpdateMode.NO_UPDATE,
      fit_columns_on_grid_load=True,
      allow_unsafe_jscode=True,
      theme="balham",
  )
