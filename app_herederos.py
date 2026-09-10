# Función para aplicar estilos visuales profesionales (Negritas, Destacados y Colores Verde/Rojo)
def aplicar_estilos_dataframe(df_styled):
  def style_cells(val, row_name, col_name):
    is_destacado = row_name in campos_destacados or row_name == "TOTAL GRUPO"
    is_total_col = (
        col_name == "Total"
        or str(col_name).startswith("Total ")
        or str(col_name).startswith("Promedio")
    )

    bg = ""
    color = ""
    weight = "bold" if (is_destacado or is_total_col) else "normal"

    if is_total_col:
      bg = "background-color: #d1fae5;"
    elif is_destacado:
      bg = "background-color: #eef2f7;"

    val_str = str(val)
    if (
        "Var." in str(col_name)
        or "pp" in val_str
        or "%" in val_str
        or "€" in val_str
    ):
      if "-" in val_str and val_str.strip() != "-":
        color = "color: #dc2626;"
        bg = "background-color: #fee2e2;"
        weight = "bold"
      elif (
          any(x in str(col_name) for x in ["Var.", "pp"])
          and val_str != "-"
          and val_str != "0,00%"
          and val_str != "0,00 pp"
          and not "-" in val_str
      ):
        color = "color: #16a34a;"
        bg = "background-color: #dcfce7;"
        weight = "bold"

    return f"{bg} {color} font-weight: {weight};"

  df_style_obj = pd.DataFrame("", index=df_styled.index, columns=df_styled.columns)
  for r_idx in df_styled.index:
    row_label = df_styled.loc[r_idx, df_styled.columns[0]]
    for c_col in df_styled.columns:
      val = df_styled.loc[r_idx, c_col]
      df_style_obj.loc[r_idx, c_col] = style_cells(val, row_label, c_col)

  return df_styled.style.apply(lambda _: df_style_obj, axis=None)
