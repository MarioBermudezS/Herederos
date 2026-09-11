import io
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode
from typing import Dict, List, Tuple
from pathlib import Path
import re



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
        padding-top: 0.25rem !important;
        padding-bottom: 0.25rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    h1 {font-size: 1.12rem !important; margin: 0 0 0.04rem 0 !important; line-height: 1.10 !important;}
    h2 {font-size: 0.98rem !important; margin: 0.04rem 0 !important; line-height: 1.10 !important;}
    h3 {font-size: 0.90rem !important; margin: 0.04rem 0 !important; line-height: 1.10 !important;}

    /* Interfaz más compacta para aprovechar toda la pantalla */
    div[data-testid="stVerticalBlock"] {gap: 0.08rem !important;}
    div[data-testid="stSidebar"] .block-container {
        padding-top: 0.25rem !important;
        padding-bottom: 0.20rem !important;
    }
    div[data-testid="stSidebar"] .stRadio,
    div[data-testid="stSidebar"] .stSelectbox,
    div[data-testid="stSidebar"] .stMultiSelect {
        margin-bottom: 0.04rem !important;
    }
    div[data-testid="stDownloadButton"] {margin-top: 0.05rem !important;}
    </style>
"""

MESES_ORDEN = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

CONCEPTOS_KPI = [
    "Ventas",
    "Consumo Ventas",
    "Variación Existencias",
    "Ajustes Existencias",
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
    """
    Convierte un número a euros con formato español.
    Los importes iguales a cero se muestran en blanco en pantalla.
    El valor numérico interno no se modifica.
    """
    try:
        if pd.isna(valor) or abs(float(valor)) < 1e-12:
            return ""
    except (TypeError, ValueError):
        pass

    return f"{valor:,.{decimales}f} €".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_variacion_pp(valor: float) -> str:
    """Formatea variaciones en puntos porcentuales."""
    return f"{valor * 100:,.2f} pp".replace(",", "X").replace(".", ",").replace("X", ".")

def calcular_ancho_columna(df: pd.DataFrame, col_name: str, min_width: int = 74) -> int:
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
    
    # Autoajuste por contenido con una reserva suficiente para que no se corten
    # cifras, signos, moneda, porcentajes ni separadores de miles.
    # Con fuente de 13 px, ~6,7 px por carácter + padding ofrece un ajuste compacto
    # pero legible incluso cuando se muestran varias columnas.
    # Ajuste algo más compacto sin cortar cifras, porcentajes ni símbolos.
    ancho = max(
        int(max_len * 4.55 + 6),
        int(len(col_name) * 4.55 + 6),
        min_width
    )

    # Ajuste compacto para aprovechar mejor el ancho de pantalla.
    return min(ancho, 150)

def obtener_archivo_datos() -> str:
    """
    Usa exclusivamente el archivo oficial BaseDatos2026.xlsx.
    No busca ni acepta copias numeradas.
    """
    archivo = Path("BaseDatos2026.xlsx")

    if not archivo.exists():
        alternativa = Path("/mnt/data/BaseDatos2026.xlsx")
        if alternativa.exists():
            archivo = alternativa

    if not archivo.exists():
        raise FileNotFoundError(
            "No se encuentra 'BaseDatos2026.xlsx'. "
            "Debe estar junto a app_herederos.py con ese nombre exacto."
        )

    return str(archivo)









def mostrar_control_base_datos():
    """
    Muestra información del Excel y permite recargar los datos
    conservando las selecciones actuales de la sesión.
    """
    try:
        archivo = Path(obtener_archivo_datos())
        stat = archivo.stat()

        from datetime import datetime
        fecha = datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S")
        tamano_mb = stat.st_size / (1024 * 1024)

        st.sidebar.markdown("---")
        st.sidebar.markdown("### Base de datos")
        st.sidebar.caption(f"Archivo: {archivo.name}")
        st.sidebar.caption(f"Actualizado: {fecha}")
        st.sidebar.caption(f"Tamaño: {tamano_mb:.2f} MB")

        if st.sidebar.button(
            "🔄 Recargar datos",
            key="btn_recargar_datos",
            use_container_width=True
        ):
            # Limpiamos únicamente la caché de datos.
            # st.session_state se conserva automáticamente en el rerun,
            # por lo que no debemos reescribir las claves de los widgets.
            st.cache_data.clear()
            st.rerun()

    except Exception as e:
        st.sidebar.warning(f"No se pudo comprobar la base de datos: {e}")

mostrar_control_base_datos()


def firma_archivo_datos() -> tuple:
    """
    Devuelve una firma del Excel para que la caché se invalide
    automáticamente cuando cambia el archivo.
    """
    archivo = Path(obtener_archivo_datos())
    stat = archivo.stat()
    return (str(archivo.resolve()), stat.st_mtime_ns, stat.st_size)



@st.cache_data(show_spinner=False)
def load_data(_firma=None) -> pd.DataFrame:
    """Carga datos del archivo Excel y normaliza campos de texto."""
    archivo_datos = obtener_archivo_datos()
    df = pd.read_excel(archivo_datos, sheet_name="BS")

    # Evita que espacios invisibles o diferencias de mayúsculas/minúsculas
    # hagan desaparecer meses, tiendas o conceptos en los filtros.
    for col in ["Mes", "Departamento", "Resultados"]:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()

    if "Mes" in df.columns:
        mapa_meses = {m.upper(): m for m in MESES_ORDEN}
        df["Mes"] = df["Mes"].apply(
            lambda x: mapa_meses.get(str(x).strip().upper(), str(x).strip())
            if pd.notna(x) else x
        )

    if "Año" in df.columns:
        df["Año"] = pd.to_numeric(df["Año"], errors="coerce")

    if "Importe D" in df.columns:
        df["Importe D"] = pd.to_numeric(df["Importe D"], errors="coerce").fillna(0.0)

    return df

@st.cache_data(show_spinner=False)
def load_tiendas_m2(_firma=None) -> Dict[str, float]:
    """Lee los metros cuadrados de la hoja Tiendas del mismo Excel."""
    try:
        archivo_datos = obtener_archivo_datos()
        t = pd.read_excel(archivo_datos, sheet_name="Tiendas")
    except Exception:
        return {}
    if "Departamento" not in t.columns or "m2" not in t.columns:
        return {}
    t["Departamento"] = t["Departamento"].astype("string").str.strip()
    t["m2"] = pd.to_numeric(t["m2"], errors="coerce")
    return {
        str(r["Departamento"]).strip(): float(r["m2"])
        for _, r in t.iterrows()
        if pd.notna(r["Departamento"]) and pd.notna(r["m2"]) and float(r["m2"]) > 0
    }




@st.cache_data(show_spinner=False)
def load_ajustes_existencias(_firma=None) -> pd.DataFrame:
    """
    Lee la hoja 'Ajustes' con columnas Año, Mes y Ajuste.

    Convención:
      Ajuste negativo -> reduce Coste/Consumo de Ventas
                      -> aumenta Margen Bruto, R.B., BAII y BAI.
    """
    archivo_datos = obtener_archivo_datos()
    columnas = ["Año", "Mes", "Ajuste"]

    try:
        aj = pd.read_excel(archivo_datos, sheet_name="Ajustes")
    except Exception:
        return pd.DataFrame(columns=columnas)

    if aj.empty:
        return pd.DataFrame(columns=columnas)

    def normalizar_nombre(c):
        return (
            str(c).strip().lower()
            .replace("á", "a").replace("é", "e").replace("í", "i")
            .replace("ó", "o").replace("ú", "u").replace("ñ", "n")
        )

    mapa = {normalizar_nombre(c): c for c in aj.columns}
    col_ano = mapa.get("ano")
    col_mes = mapa.get("mes")
    col_ajuste = next(
        (
            mapa.get(k)
            for k in [
                "ajuste",
                "ajustes",
                "ajuste existencias",
                "ajustes existencias",
                "variacion existencias",
                "variacion de existencias",
            ]
            if mapa.get(k) is not None
        ),
        None,
    )

    if col_ano is None or col_mes is None or col_ajuste is None:
        return pd.DataFrame(columns=columnas)

    out = aj[[col_ano, col_mes, col_ajuste]].copy()
    out.columns = columnas
    out["Año"] = pd.to_numeric(out["Año"], errors="coerce")
    out["Ajuste"] = pd.to_numeric(out["Ajuste"], errors="coerce")

    mapa_meses = {m.upper(): m for m in MESES_ORDEN}
    out["Mes"] = out["Mes"].apply(
        lambda x: mapa_meses.get(str(x).strip().upper(), str(x).strip())
        if pd.notna(x) else x
    )

    out = out.dropna(subset=["Año", "Mes", "Ajuste"]).copy()
    out["Año"] = out["Año"].astype(int)

    return out.groupby(["Año", "Mes"], as_index=False)["Ajuste"].sum()


def obtener_ajuste_existencias(ano: int, meses: List[str]) -> float:
    ajustes = load_ajustes_existencias(firma_archivo_datos())
    if ajustes.empty or not meses:
        return 0.0

    filtro = ajustes[
        (ajustes["Año"] == int(ano))
        & (ajustes["Mes"].isin(list(meses)))
    ]
    if filtro.empty:
        return 0.0

    return float(
        pd.to_numeric(filtro["Ajuste"], errors="coerce").fillna(0).sum()
    )


@st.cache_data(show_spinner=False)
def load_margenes_totales_acumulados(_firma=None) -> pd.DataFrame:
    """
    Lee de forma estricta la hoja MargenesTotalesAcumulados del ERP.

    Para cada bloque anual:
      - localiza el año (2024, 2025, 2026, etc.)
      - localiza la fila de cabeceras con Enero...Diciembre
      - localiza la fila TOTAL/TOTALES inmediatamente asociada
      - relaciona cada mes con su valor exacto

    Devuelve:
        Año | Mes | Margen Acumulado Total
    """
    archivo_datos = obtener_archivo_datos()

    try:
        raw = pd.read_excel(
            archivo_datos,
            sheet_name="MargenesTotalesAcumulados",
            header=None,
        )
    except Exception:
        return pd.DataFrame(
            columns=["Año", "Mes", "Margen Acumulado Total"]
        )

    registros = []
    mapa_meses = {m.upper(): m for m in MESES_ORDEN}

    i = 0
    while i < len(raw):
        ano_actual = None

        # El año puede estar en cualquiera de las primeras columnas.
        for j in range(min(3, raw.shape[1])):
            valor = raw.iat[i, j]
            try:
                ano_posible = int(float(valor))
                if 2000 <= ano_posible <= 2100:
                    ano_actual = ano_posible
                    break
            except (TypeError, ValueError):
                pass

        if ano_actual is None:
            i += 1
            continue

        # Buscar la fila de cabeceras de meses dentro del bloque de ese año.
        fila_cabecera = None
        columnas_mes = {}

        for h in range(i + 1, min(i + 6, len(raw))):
            encontrados = {}
            for j in range(raw.shape[1]):
                valor = raw.iat[h, j]
                if pd.isna(valor):
                    continue
                mes = mapa_meses.get(str(valor).strip().upper())
                if mes:
                    encontrados[j] = mes

            if encontrados:
                fila_cabecera = h
                columnas_mes = encontrados
                break

        if fila_cabecera is None:
            i += 1
            continue

        # Buscar la fila TOTAL/TOTALES asociada a esa cabecera,
        # sin cruzar al siguiente bloque anual.
        fila_totales = None
        for t in range(fila_cabecera + 1, min(fila_cabecera + 15, len(raw))):
            # Si aparece otro año antes de Totales, abandonar este bloque.
            nuevo_ano = False
            for j in range(min(3, raw.shape[1])):
                valor = raw.iat[t, j]
                try:
                    ano_posible = int(float(valor))
                    if 2000 <= ano_posible <= 2100:
                        nuevo_ano = True
                        break
                except (TypeError, ValueError):
                    pass
            if nuevo_ano:
                break

            primera = raw.iat[t, 0] if raw.shape[1] else None
            primera_txt = "" if pd.isna(primera) else str(primera).strip().upper()

            if primera_txt in {"TOTAL", "TOTALES"}:
                fila_totales = t
                break

        if fila_totales is not None:
            for col, mes in columnas_mes.items():
                valor = pd.to_numeric(
                    raw.iat[fila_totales, col],
                    errors="coerce",
                )
                if pd.notna(valor):
                    margen = float(valor)

                    # El ERP normalmente guarda 42,72 % como 0,4272.
                    # Si viniera como 42,72, se normaliza.
                    if abs(margen) > 1.0:
                        margen = margen / 100.0

                    registros.append(
                        {
                            "Año": ano_actual,
                            "Mes": mes,
                            "Margen Acumulado Total": margen,
                        }
                    )

        i = max(i + 1, (fila_totales + 1) if fila_totales is not None else i + 1)

    df_margenes = pd.DataFrame(registros)

    if df_margenes.empty:
        return pd.DataFrame(
            columns=["Año", "Mes", "Margen Acumulado Total"]
        )

    # Si hubiera duplicados accidentales, conservar el último valor de cada mes/año.
    df_margenes = (
        df_margenes
        .drop_duplicates(subset=["Año", "Mes"], keep="last")
        .reset_index(drop=True)
    )

    return df_margenes


@st.cache_data(show_spinner=False)
def load_inventario(_firma=None) -> pd.DataFrame:
    """
    Lee la hoja Inventario organizada por bloques de año.

    Formato esperado:
      fila con el año
      fila con Departamento + Enero...Diciembre
      filas de tiendas
      fila de total sin nombre

    Devuelve:
      Año | Mes | Departamento | Inventario
    """
    archivo_datos = obtener_archivo_datos()

    try:
        raw = pd.read_excel(archivo_datos, sheet_name="Inventario", header=None)
    except Exception:
        return pd.DataFrame(
            columns=["Año", "Mes", "Departamento", "Inventario"]
        )

    registros = []
    ano_actual = None
    mapa_meses = {m.upper(): m for m in MESES_ORDEN}

    for i in range(len(raw)):
        primera = raw.iat[i, 0] if raw.shape[1] else None

        # Inicio de bloque anual.
        try:
            ano_posible = int(float(primera))
            if 2000 <= ano_posible <= 2100:
                ano_actual = ano_posible
                continue
        except (TypeError, ValueError):
            pass

        if ano_actual is None:
            continue

        if str(primera).strip().upper() != "DEPARTAMENTO":
            continue

        cabeceras = []
        for j in range(1, raw.shape[1]):
            valor = raw.iat[i, j]
            cabeceras.append(
                mapa_meses.get(str(valor).strip().upper())
            )

        k = i + 1
        while k < len(raw):
            dep = raw.iat[k, 0]

            # Si aparece otro año, termina el bloque actual.
            try:
                ano_siguiente = int(float(dep))
                if 2000 <= ano_siguiente <= 2100:
                    break
            except (TypeError, ValueError):
                pass

            dep_txt = "" if pd.isna(dep) else str(dep).strip()

            # La fila total del Excel no se importa; Python recalcula siempre.
            if not dep_txt or dep_txt.upper() in {"TOTAL", "DEPARTAMENTO"}:
                k += 1
                continue

            for j, mes in enumerate(cabeceras, start=1):
                if not mes or j >= raw.shape[1]:
                    continue

                valor = pd.to_numeric(raw.iat[k, j], errors="coerce")
                if pd.notna(valor):
                    registros.append(
                        {
                            "Año": ano_actual,
                            "Mes": mes,
                            "Departamento": dep_txt,
                            "Inventario": float(valor),
                        }
                    )
            k += 1

    inventario = pd.DataFrame(registros)
    if inventario.empty:
        return pd.DataFrame(
            columns=["Año", "Mes", "Departamento", "Inventario"]
        )

    inventario["Departamento"] = (
        inventario["Departamento"].astype("string").str.strip()
    )
    inventario["Año"] = pd.to_numeric(inventario["Año"], errors="coerce")
    inventario["Inventario"] = pd.to_numeric(
        inventario["Inventario"], errors="coerce"
    )

    return inventario


def obtener_ultimos_12_periodos(
    ano: int, mes: str
) -> List[Tuple[int, str]]:
    """
    Devuelve 12 meses terminando en el mes analizado.
    Ejemplo: julio-2026 => agosto-2025 ... julio-2026.
    """
    if mes not in MESES_ORDEN:
        return []

    indice_absoluto = int(ano) * 12 + MESES_ORDEN.index(mes)
    periodos = []

    for desplazamiento in range(11, -1, -1):
        idx = indice_absoluto - desplazamiento
        periodos.append((idx // 12, MESES_ORDEN[idx % 12]))

    return periodos


def calcular_cobertura_inventario(
    df_datos: pd.DataFrame,
    tienda: str,
    ano: int,
    mes: str,
    inventario_final: float,
) -> Dict[str, object]:
    """
    Método de rotación/cobertura:

    - Media mensual de ventas de los últimos 12 meses, incluido el analizado.
    - Margen bruto acumulado de la tienda desde enero hasta ese mes.
    - Venta media a coste = Venta media 12M * (1 - margen acumulado).
    - Meses de stock = Inventario final / Venta media mensual a coste.

    Si no existen 12 meses completos de ventas, no se calcula.
    """
    mensaje_12m = "Imposible calcular media de ventas últimos 12 meses"

    periodos = obtener_ultimos_12_periodos(ano, mes)
    if len(periodos) != 12:
        return {
            "venta_media_12m": None,
            "margen_acumulado": None,
            "venta_media_coste": None,
            "meses_stock": None,
            "estado": mensaje_12m,
        }

    ventas_mensuales = []

    for a, m in periodos:
        df_mes = df_datos[
            (df_datos["Año"] == a)
            & (df_datos["Mes"] == m)
            & (
                df_datos["Departamento"].astype(str).str.strip()
                == str(tienda).strip()
            )
        ]

        ventas_rows = df_mes[df_mes["Resultados"] == "Ventas"]

        # Debe existir información de ventas para cada uno de los 12 meses.
        if ventas_rows.empty:
            return {
                "venta_media_12m": None,
                "margen_acumulado": None,
                "venta_media_coste": None,
                "meses_stock": None,
                "estado": mensaje_12m,
            }

        ventas_mes = float(
            pd.to_numeric(
                ventas_rows["Importe D"], errors="coerce"
            ).fillna(0).sum()
        )
        ventas_mensuales.append(ventas_mes)

    venta_media_12m = sum(ventas_mensuales) / 12.0

    # Margen acumulado del ejercicio de la tienda hasta el mes analizado.
    meses_acumulados = MESES_ORDEN[: MESES_ORDEN.index(mes) + 1]
    df_acumulado = obtener_filtro_datos(
        df_datos, ano, meses_acumulados, [tienda]
    )
    resultados_acumulados = calcular_resultados(df_acumulado)
    margen_acumulado = float(
        resultados_acumulados.get("R. B.", 0.0)
    )

    venta_media_coste = venta_media_12m * (1.0 - margen_acumulado)

    if abs(venta_media_coste) < 1e-12:
        return {
            "venta_media_12m": venta_media_12m,
            "margen_acumulado": margen_acumulado,
            "venta_media_coste": venta_media_coste,
            "meses_stock": None,
            "estado": "Imposible calcular: venta media a coste igual a cero",
        }

    return {
        "venta_media_12m": venta_media_12m,
        "margen_acumulado": margen_acumulado,
        "venta_media_coste": venta_media_coste,
        "meses_stock": float(inventario_final) / venta_media_coste,
        "estado": "",
    }


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


def seleccion_es_total_empresa(
    tiendas: List[str],
    departamentos_sin_general: List[str],
) -> bool:
    """Comprueba si están seleccionadas todas las tiendas operativas."""
    seleccion = {
        str(x).strip()
        for x in (tiendas or [])
        if str(x).strip().upper() != "GENERAL"
    }
    total = {
        str(x).strip()
        for x in (departamentos_sin_general or [])
        if str(x).strip().upper() != "GENERAL"
    }
    return bool(total) and seleccion == total


def calcular_resultados_seleccion(
    df_filtrado: pd.DataFrame,
    ano: int,
    meses: List[str],
    tiendas: List[str],
    departamentos_sin_general: List[str],
) -> Dict[str, float]:
    """
    Aplica Ajustes Existencias únicamente cuando la selección es TOTAL empresa.
    """
    ajuste = 0.0
    if seleccion_es_total_empresa(tiendas, departamentos_sin_general):
        ajuste = obtener_ajuste_existencias(ano, meses)

    return calcular_resultados(
        df_filtrado,
        ajuste_existencias=ajuste,
    )


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

def _periodo_anterior(ano: int, mes: str) -> Tuple[int, str]:
    idx = MESES_ORDEN.index(mes)
    if idx == 0:
        return int(ano) - 1, "Diciembre"
    return int(ano), MESES_ORDEN[idx - 1]


def obtener_inventario_total_mes(ano: int, mes: str) -> float:
    """Inventario total empresa del cierre del mes, incluido General."""
    inv = load_inventario(firma_archivo_datos())
    if inv.empty:
        return 0.0
    f = inv[(inv["Año"] == int(ano)) & (inv["Mes"] == mes)]
    if f.empty:
        return 0.0
    return float(pd.to_numeric(f["Inventario"], errors="coerce").fillna(0).sum())


def obtener_variacion_existencias_mes(ano: int, mes: str) -> float:
    """
    Variación de existencias = Inventario inicial - Inventario final.
    Ejemplo julio 2026: 3.307.283,17 - 3.510.319,10 = -203.035,93.
    """
    ano_ant, mes_ant = _periodo_anterior(ano, mes)
    inv_ini = obtener_inventario_total_mes(ano_ant, mes_ant)
    inv_fin = obtener_inventario_total_mes(ano, mes)
    if inv_ini == 0.0 or inv_fin == 0.0:
        return 0.0
    return inv_ini - inv_fin



def calcular_capitulo_financiero(df_periodo: pd.DataFrame) -> tuple[float, float]:
    """
    Calcula los capítulos financieros por naturaleza contable de la cuenta.

    Gastos financieros:
      - 626 Servicios bancarios y similares
      - cuentas 66x (p. ej. 662, 669)

    Ingresos financieros:
      - cuentas 76x (p. ej. 760, 763, 769)

    Se conserva SIEMPRE el signo original de 'Importe D'.
    """
    if df_periodo is None or df_periodo.empty:
        return 0.0, 0.0

    tmp = df_periodo.copy()

    if "Cuenta" not in tmp.columns or "Importe D" not in tmp.columns:
        return 0.0, 0.0

    tmp["Cuenta_txt"] = (
        tmp["Cuenta"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )
    tmp["Importe_num"] = pd.to_numeric(
        tmp["Importe D"], errors="coerce"
    ).fillna(0.0)

    es_gasto_financiero = (
        tmp["Cuenta_txt"].eq("626")
        | tmp["Cuenta_txt"].str.startswith("66", na=False)
    )

    es_ingreso_financiero = tmp["Cuenta_txt"].str.startswith("76", na=False)

    gastos_financieros = float(
        tmp.loc[es_gasto_financiero, "Importe_num"].sum()
    )
    ingresos_financieros = float(
        tmp.loc[es_ingreso_financiero, "Importe_num"].sum()
    )

    return gastos_financieros, ingresos_financieros


def calcular_resultados_total_empresa(
    df_datos: pd.DataFrame,
    ano: int,
    meses: List[str],
) -> Dict[str, float]:
    """
    Cuenta de Resultados TOTAL empresa, incluyendo el departamento General.

    Criterio acordado:
      Coste Ventas = Consumo Ventas + Variación Existencias + Ajustes Existencias
      Variación Existencias = Inventario inicial - Inventario final
      Margen Bruto = Ventas - Coste Ventas

    Los importes mensuales se calculan mes a mes y después se acumulan.
    """
    meses_validos = [m for m in MESES_ORDEN if m in list(meses or [])]
    if not meses_validos:
        return {c: 0.0 for c in CONCEPTOS_KPI}

    df_periodo = obtener_filtro_datos(df_datos, int(ano), meses_validos, None)
    if df_periodo.empty:
        return {c: 0.0 for c in CONCEPTOS_KPI}

    resumen = df_periodo.groupby("Resultados")["Importe D"].sum().to_dict()

    def get_v(cat):
        return float(resumen.get(cat, 0.0) or 0.0)

    ventas = get_v("Ventas")
    consumo_ventas = get_v("Consumo Ventas")
    variacion_existencias = sum(
        obtener_variacion_existencias_mes(int(ano), mes) for mes in meses_validos
    )
    ajustes_existencias = obtener_ajuste_existencias(int(ano), meses_validos)

    coste_ventas = consumo_ventas + variacion_existencias + ajustes_existencias
    margen_bruto = ventas - coste_ventas
    r_bruta = (margen_bruto / ventas) if ventas != 0 else 0.0

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

    baii = ingresos_operativos - gastos_estructura

    # IMPORTANTE: conservar SIEMPRE el signo contable original del Excel.
    # En las cuentas de ingresos (grupo 7), los importes pueden venir negativos.
    # Por ejemplo:
    #   Gastos Financieros      +10.000
    #   Ingresos Financieros     -2.000
    #   RDO. FINANCIERO           8.000  -> resta 8.000 al BAI
    #
    # Del mismo modo, Resultados Extraordinarios se toma tal cual del Excel:
    #   positivo -> resta al BAI
    #   negativo -> al restarlo, aumenta el BAI (es ingreso neto extraordinario)
    gastos_financieros, ingresos_financieros = calcular_capitulo_financiero(
        df_periodo
    )
    rdo_financiero = gastos_financieros + ingresos_financieros

    resultados_extraordinarios = get_v("Resultados Extraordinarios")
    bai = baii - rdo_financiero - resultados_extraordinarios

    return {
        "Ventas": ventas,
        "Consumo Ventas": consumo_ventas,
        "Variación Existencias": variacion_existencias,
        "Ajustes Existencias": ajustes_existencias,
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


def calcular_resultados(
    df_filtrado: pd.DataFrame,
    ajuste_existencias: float = 0.0,
) -> Dict[str, float]:
    """Calcula la cuenta de resultados, incluyendo ajustes globales opcionales."""
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
    
    # Ajuste de existencias del TOTAL empresa.
    # Ajuste negativo -> reduce Coste/Consumo y aumenta Margen y Resultado.
    ajuste_existencias = float(ajuste_existencias or 0.0)
    margen_bruto = margen_bruto - ajuste_existencias

    base_rb = total_ventas_calc if total_ventas_calc != 0 else ventas
    r_bruta = (margen_bruto / base_rb) if base_rb != 0 else 0.0

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
    
    # Resultado financiero y extraordinario.
    # CONSERVAR el signo contable que viene del Excel:
    # - los ingresos financieros negativos reducen el saldo financiero y mejoran BAI;
    # - un resultado extraordinario negativo, al restarse, aumenta el BAI.
    gastos_financieros, ingresos_financieros = calcular_capitulo_financiero(
        df_filtrado
    )
    rdo_financiero = gastos_financieros + ingresos_financieros

    resultados_extraordinarios = get_v("Resultados Extraordinarios")
    bai = baii - rdo_financiero - resultados_extraordinarios
    
    return {
        "Ventas": ventas,
        "Consumo Ventas": 0.0,
        "Variación Existencias": 0.0,
        "Coste Ventas": coste_ventas,
        "Ajustes Existencias": ajuste_existencias,
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

def _safe_storage_key(clave: str) -> str:
    return "herederos_grid_" + "".join(
        c if c.isalnum() or c in "_-" else "_" for c in str(clave)
    )


def estado_columnas_js(clave: str, autoajustar_todas: bool = False) -> JsCode:
    """Guarda/restaura ancho y orden de columnas en el navegador."""
    storage_key = _safe_storage_key(clave)
    autoajuste_js = "true" if autoajustar_todas else "false"

    return JsCode(
        f"""
        function(params) {{
            const key = {storage_key!r};
            const autoajustarTodas = {autoajuste_js};

            function guardar() {{
                try {{
                    window.localStorage.setItem(
                        key,
                        JSON.stringify(params.api.getColumnState())
                    );
                }} catch (e) {{}}
            }}

            let hayEstadoGuardado = false;
            try {{
                const saved = window.localStorage.getItem(key);
                hayEstadoGuardado = !!saved;
                if (saved && !autoajustarTodas) {{
                    params.api.applyColumnState({{
                        state: JSON.parse(saved),
                        applyOrder: true
                    }});
                }}
            }} catch (e) {{}}

            // Primera vez para esta combinación de columnas: autoajuste inicial.
            // Con muchas columnas se prioriza que entren en el ancho disponible.
            if (!autoajustarTodas) {{
                try {{
                    let numeroColumnas = 0;
                    if (params.api && params.api.getColumnState) {{
                        numeroColumnas = params.api.getColumnState().length;
                    }}

                    if (
                        numeroColumnas >= 8 &&
                        params.api &&
                        params.api.sizeColumnsToFit
                    ) {{
                        setTimeout(function() {{
                            try {{
                                params.api.sizeColumnsToFit();
                                guardar();
                            }} catch (e) {{}}
                        }}, 120);
                    }} else {{
                        if (params.api && params.api.autoSizeAllColumns) {{
                            params.api.autoSizeAllColumns(true);
                        }} else if (
                            params.columnApi &&
                            params.columnApi.autoSizeAllColumns
                        ) {{
                            params.columnApi.autoSizeAllColumns(true);
                        }}
                        setTimeout(guardar, 220);
                    }}
                }} catch (e) {{}}
            }}

            // Autoajuste de todas las columnas cuando se pulsa el botón.
            if (autoajustarTodas) {{
                try {{
                    if (params.api && params.api.autoSizeAllColumns) {{
                        params.api.autoSizeAllColumns(true);
                        setTimeout(function() {{
                            try {{
                                const state = params.api.getColumnState();
                                const compact = state.map(function(col, index) {{
                                    if (col.width && col.width > 70) {{
                                        if (index === 0) {{
                                            // Primera columna: conservar más ancho para conceptos.
                                            col.width = Math.max(140, Math.round(col.width * 0.95));
                                        }} else {{
                                            col.width = Math.max(52, Math.round(col.width * 0.84));
                                        }}
                                    }}
                                    return col;
                                }});
                                params.api.applyColumnState({{
                                    state: compact,
                                    applyOrder: true
                                }});
                            }} catch (e) {{}}
                        }}, 80);
                    }} else if (params.columnApi && params.columnApi.autoSizeAllColumns) {{
                        params.columnApi.autoSizeAllColumns(false);
                    }} else {{
                        const cols = [];
                        if (params.api && params.api.getColumns) {{
                            params.api.getColumns().forEach(function(col) {{
                                cols.push(col.getColId());
                            }});
                        }}
                        if (cols.length && params.columnApi && params.columnApi.autoSizeColumns) {{
                            params.columnApi.autoSizeColumns(cols, false);
                        }}
                    }}
                    setTimeout(guardar, 250);
                }} catch (e) {{}}
            }}

            // Los cambios manuales y el doble clic en el separador
            // también terminan guardándose automáticamente.
            params.api.addEventListener('columnResized', function(e) {{
                if (e.finished) guardar();
            }});
            params.api.addEventListener('columnMoved', function(e) {{
                if (e.finished) guardar();
            }});
        }}
        """
    )

def es_columna_resumen(nombre_columna: str) -> bool:
    """
    Identifica únicamente columnas de resumen reales.
    Ejemplos que SÍ: Total, Total Grupo, Promedio, Promedio Acumulado, Acumulado.
    Ejemplos que NO: Total 2026, Total 2025.
    """
    nombre = str(nombre_columna).strip().upper()
    return nombre in {
        "TOTAL",
        "TOTAL GRUPO",
        "PROMEDIO",
        "PROMEDIO ACUMULADO",
        "ACUMULADO",
    }


def render_aggrid_table(
    df_display: pd.DataFrame,
    modo: str = "auto",
    df_numericos: pd.DataFrame = None,
    resaltar_kpi_tiendas: bool = False,
    columnas_comparar: List[str] = None,
    resaltar_rb_tiendas: bool = False,
    resaltar_extremos_filas: bool = False,
    columnas_extremos: List[str] = None,
    altura_fila: int = 32,
    altura_cabecera: int = 34,
    clave_preferencias: str = "tabla_general",
) -> None:
    """
    Renderiza la tabla completa.

    - Muestra todas las filas sin scroll vertical interno.
    - Autoajusta cada columna según cabecera y contenido.
    - Resalta con negrita y sombreado las líneas principales.
    - Opcionalmente, en Informe KPI multi-tienda:
        * Gastos/ratios: menor % verde y mayor % rojo.
        * Ingresos: mayor % verde y menor % rojo.
      La columna Total queda fuera de la comparación.
    """
    if df_display.empty:
        st.info("No hay datos para mostrar con los filtros seleccionados.")
        return

    # Diseño compacto: más filas visibles sin perder legibilidad.
    altura_fila = min(int(altura_fila), 27)
    altura_cabecera = min(int(altura_cabecera), 30)

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
        "TOTAL",
        "Total",
    }

    filas_js = ",".join(repr(x) for x in sorted(filas_negrita))
    get_row_style = JsCode(
        f"""
        function(params) {{
            let valor = '';
            if (params.data) {{
                if (params.data.Resultados !== undefined && params.data.Resultados !== null) {{
                    valor = String(params.data.Resultados);
                }} else if (params.data.Tienda !== undefined && params.data.Tienda !== null) {{
                    valor = String(params.data.Tienda);
                }} else {{
                    const claves = Object.keys(params.data);
                    if (claves.length > 0 && params.data[claves[0]] !== undefined && params.data[claves[0]] !== null) {{
                        valor = String(params.data[claves[0]]);
                    }}
                }}
            }}

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

    # Filas consideradas "ingresos": en ellas el mayor porcentaje es mejor.
    conceptos_ingresos = {
        "Ventas",
        "Otros Ingresos",
        "Ingresos Operativos",
        "Ingresos Financieros",
    }

    # Para cada columna/tienda, preparar qué filas deben ir verde o rojo.
    verdes_por_columna = {}
    rojos_por_columna = {}

    if (
        resaltar_kpi_tiendas
        and df_numericos is not None
        and columnas_comparar
        and len(columnas_comparar) > 1
    ):
        columnas_validas = [
            c for c in columnas_comparar
            if c in df_numericos.columns and c in df_display.columns
        ]

        if len(columnas_validas) > 1:
            for c in columnas_validas:
                verdes_por_columna[c] = set()
                rojos_por_columna[c] = set()

            for _, fila in df_numericos.iterrows():
                concepto = str(fila["Resultados"])
                valores = {}

                for c in columnas_validas:
                    try:
                        valor = float(fila[c])
                        if pd.notna(valor):
                            valores[c] = valor
                    except (TypeError, ValueError):
                        pass

                if len(valores) < 2:
                    continue

                minimo = min(valores.values())
                maximo = max(valores.values())

                # Si todas las tiendas tienen el mismo valor no se colorea ninguna.
                if minimo == maximo:
                    continue

                es_ingreso = concepto in conceptos_ingresos

                for c, valor in valores.items():
                    if es_ingreso:
                        if valor == maximo:
                            verdes_por_columna[c].add(concepto)
                        if valor == minimo:
                            rojos_por_columna[c].add(concepto)
                    else:
                        if valor == minimo:
                            verdes_por_columna[c].add(concepto)
                        if valor == maximo:
                            rojos_por_columna[c].add(concepto)


    # Resaltado específico del análisis R.B.:
    # entre tiendas, el valor más alto va en verde y el más bajo en rojo.
    rb_verdes = {}
    rb_rojos = {}
    if (
        resaltar_rb_tiendas
        and df_numericos is not None
        and columnas_comparar
        and len(columnas_comparar) > 1
    ):
        # En R.B. las tiendas están en filas y los periodos/medidas en columnas.
        tiendas_validas = set(columnas_comparar)
        filas_tiendas = df_numericos[df_numericos["Resultados"].isin(tiendas_validas)]

        for col in df_numericos.columns:
            if col == "Resultados":
                continue

            # Totales, promedios y acumulados no participan en máximos/mínimos.
            col_norm = str(col).strip().upper()
            if es_columna_resumen(col):
                continue

            valores = {}
            for _, fila in filas_tiendas.iterrows():
                try:
                    valor = float(fila[col])
                    if pd.notna(valor):
                        valores[str(fila["Resultados"])] = valor
                except (TypeError, ValueError):
                    pass

            if len(valores) < 2:
                continue

            minimo = min(valores.values())
            maximo = max(valores.values())

            if minimo == maximo:
                continue

            rb_verdes[col] = {tienda for tienda, valor in valores.items() if valor == maximo}
            rb_rojos[col] = {tienda for tienda, valor in valores.items() if valor == minimo}

    # Resaltado opcional de mejor/peor valor por columna entre filas.
    # Se excluye la fila TOTAL para no mezclar el agregado con las tiendas.
    extremos_mejor = {}
    extremos_peor = {}
    if (
        resaltar_extremos_filas
        and df_numericos is not None
        and columnas_extremos
    ):
        for col_ext in columnas_extremos:
            if col_ext not in df_numericos.columns or col_ext not in df_display.columns:
                continue
            serie = pd.to_numeric(df_numericos[col_ext], errors="coerce")
            mascara = pd.Series(True, index=df_numericos.index)
            if "Tienda" in df_numericos.columns:
                mascara = (
                    df_numericos["Tienda"].astype(str).str.strip().str.upper() != "TOTAL"
                )
            # Para mejor/peor se excluyen valores vacíos, NaN y cero.
            # Así una tienda sin dato real no aparece artificialmente como "peor".
            validos = serie[mascara & serie.notna() & serie.ne(0)]
            if validos.empty:
                continue
            maximo = validos.max()
            minimo = validos.min()
            extremos_mejor[col_ext] = set(
                df_numericos.loc[mascara & serie.eq(maximo), "Tienda"].astype(str)
            )
            extremos_peor[col_ext] = set(
                df_numericos.loc[mascara & serie.eq(minimo), "Tienda"].astype(str)
            )

    gb = GridOptionsBuilder.from_dataframe(df_display)
    # El usuario puede redimensionar arrastrando el borde de la cabecera
    # y reordenar columnas arrastrando la propia cabecera.
    gb.configure_default_column(
        resizable=True,
        sortable=False,
        filter=False,
        suppressMovable=False,
    )
    gb.configure_default_column(
        resizable=True,
        filterable=False,
        sortable=False,
        editable=False,
        suppressMenu=True,
        wrapText=False,
        autoHeight=False,
    )

    # Estilo numérico normal: negativos en rojo.
    estilo_numerico_js = JsCode(
        r"""
        function(params) {
            const raw = params.value;
            if (raw === null || raw === undefined) {
                return {'textAlign': 'right'};
            }

            const texto = String(raw).trim();
            const esNegativo = texto.startsWith('-') || /^\(.*\)$/.test(texto);

            if (esNegativo) {
                return {
                    'textAlign': 'right',
                    'color': '#d00000'
                };
            }
            return {'textAlign': 'right'};
        }
        """
    )

    for i, col in enumerate(df_display.columns):
        if i == 0:
            ancho = calcular_ancho_columna(df_display, col, 118)
            # Concepto se mantiene prácticamente igual, solo un poco más compacto.
            ancho_primera = max(ancho, 146)
            gb.configure_column(
                col,
                width=ancho_primera,
                minWidth=140,
                maxWidth=225,
                cellStyle={"textAlign": "left", "fontWeight": "600"},
            )
        else:
            ancho = calcular_ancho_columna(df_display, col, 66)

            col_norm = str(col).strip().upper()
            es_columna_total = es_columna_resumen(col)

            # En Eficiencia de Inventario: mejor valor verde y peor naranja.
            if col in extremos_mejor or col in extremos_peor:
                mejores = ",".join(repr(x) for x in sorted(extremos_mejor.get(col, set())))
                peores = ",".join(repr(x) for x in sorted(extremos_peor.get(col, set())))
                estilo_extremos_js = JsCode(
                    f"""
                    function(params) {{
                        const tienda = params.data && params.data.Tienda !== undefined
                            ? String(params.data.Tienda)
                            : '';
                        const mejores = [{mejores}];
                        const peores = [{peores}];

                        let estilo = {{'textAlign': 'right'}};

                        if (mejores.includes(tienda)) {{
                            estilo['backgroundColor'] = '#d9ead3';
                            estilo['fontWeight'] = '700';
                        }} else if (peores.includes(tienda)) {{
                            estilo['backgroundColor'] = '#fce5cd';
                            estilo['fontWeight'] = '700';
                        }}

                        const raw = params.value;
                        const texto = raw === null || raw === undefined ? '' : String(raw).trim();
                        if (texto.startsWith('-') || /^\\(.*\\)$/.test(texto)) {{
                            estilo['color'] = '#d00000';
                        }}
                        return estilo;
                    }}
                    """
                )
                cell_style = estilo_extremos_js

            # Las columnas de total/promedio/acumulado tienen sombreado propio
            # y no usan el semáforo verde/rojo comparativo.
            elif es_columna_total:
                estilo_total_js = JsCode(
                    r"""
                    function(params) {
                        const raw = params.value;
                        const texto = raw === null || raw === undefined ? '' : String(raw).trim();

                        let estilo = {
                            'textAlign': 'right',
                            'backgroundColor': '#fff7cc',
                            'fontWeight': '700'
                        };

                        if (texto.startsWith('-') || /^\(.*\)$/.test(texto)) {
                            estilo['color'] = '#d00000';
                        }

                        return estilo;
                    }
                    """
                )
                cell_style = estilo_total_js

            # Si esta columna participa en comparación KPI o R.B., añadir sombreado.
            elif (
                col in verdes_por_columna
                or col in rojos_por_columna
                or col in rb_verdes
                or col in rb_rojos
            ):
                filas_verdes_set = set(verdes_por_columna.get(col, set())) | set(rb_verdes.get(col, set()))
                filas_rojas_set = set(rojos_por_columna.get(col, set())) | set(rb_rojos.get(col, set()))
                filas_verdes = ",".join(repr(x) for x in sorted(filas_verdes_set))
                filas_rojas = ",".join(repr(x) for x in sorted(filas_rojas_set))

                estilo_kpi_js = JsCode(
                    f"""
                    function(params) {{
                        const concepto = params.data && params.data.Resultados
                            ? String(params.data.Resultados)
                            : '';
                        const verdes = [{filas_verdes}];
                        const rojos = [{filas_rojas}];

                        const raw = params.value;
                        const texto = raw === null || raw === undefined ? '' : String(raw).trim();
                        const esNegativo = texto.startsWith('-') || /^\\(.*\\)$/.test(texto);

                        let estilo = {{
                            'textAlign': 'right'
                        }};

                        if (verdes.includes(concepto)) {{
                            estilo['backgroundColor'] = '#d9ead3';
                            estilo['fontWeight'] = '700';
                        }} else if (rojos.includes(concepto)) {{
                            estilo['backgroundColor'] = '#f4cccc';
                            estilo['fontWeight'] = '700';
                        }}

                        if (esNegativo) {{
                            estilo['color'] = '#d00000';
                        }}

                        return estilo;
                    }}
                    """
                )
                cell_style = estilo_kpi_js
            else:
                cell_style = estilo_numerico_js

            gb.configure_column(
                col,
                width=ancho,
                minWidth=52,
                maxWidth=148,
                cellStyle=cell_style,
            )

    gb.configure_grid_options(
        domLayout="normal",
        suppressRowClickSelection=True,
        rowHeight=altura_fila,
        headerHeight=altura_cabecera,
        getRowStyle=get_row_style,
        suppressHorizontalScroll=False,
    )

    altura_tabla = altura_cabecera + (len(df_display) * altura_fila)

    # Cada combinación distinta de columnas tiene su propia configuración.
    # Primera vez: autoajuste. Si se vuelve a la misma combinación: recuerda ajustes.
    firma_columnas = "__".join(str(c) for c in df_display.columns)

    # La preferencia de columnas pertenece a ESTA consulta concreta.
    # Así, al cambiar año o meses, la nueva información arranca siempre
    # con autoajuste inicial y no hereda anchos de una consulta anterior.
    ano_actual = globals().get("ano", "")
    meses_actuales = globals().get("meses_sel", [])
    firma_consulta = f"{ano_actual}__{'__'.join(str(m) for m in meses_actuales)}"

    clave_preferencias_efectiva = (
        f"{clave_preferencias}__{firma_consulta}__{firma_columnas}"
    )

    boton_autoajuste = st.button(
        "↔ Autoajustar todas las columnas",
        key=f"autoajustar_todas_{clave_preferencias_efectiva}",
        help="Ajusta automáticamente todas las columnas según su cabecera y contenido.",
    )

    token_key = f"token_autoajuste_{clave_preferencias_efectiva}"
    if token_key not in st.session_state:
        st.session_state[token_key] = 0

    if boton_autoajuste:
        st.session_state[token_key] += 1

    autoajustar_todas = boton_autoajuste

    gb.configure_grid_options(
        onGridReady=estado_columnas_js(
            clave_preferencias_efectiva,
            autoajustar_todas=autoajustar_todas,
        )
    )

    AgGrid(
        df_display,
        gridOptions=gb.build(),
        key=f"grid_{clave_preferencias_efectiva}_{st.session_state[token_key]}",
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        theme="balham",
        height=altura_tabla,
        custom_css={
            ".ag-cell": {
                "font-size": "12.5px",
                "line-height": "24px",
                "padding-left": "2px",
                "padding-right": "2px",
                "border-right": "1px solid #c9ced3",
            },
            ".ag-header-cell": {
                "font-size": "12.5px",
                "font-weight": "600",
                "padding-left": "2px",
                "padding-right": "2px",
                "border-right": "1px solid #b8bec5",
            },
            ".ag-header-cell-label": {
                "justify-content": "center !important",
                "text-align": "center !important",
                "width": "100%",
            },
            ".ag-header-cell-text": {
                "text-align": "center !important",
                "width": "100%",
            },
        },
    )

def render_aggrid_rb_horizontal(
    df_display: pd.DataFrame,
    df_numericos: pd.DataFrame,
    tiendas: List[str],
    altura_fila: int = 36,
    altura_cabecera: int = 38,
    clave_preferencias: str = "rb_meses_filas",
) -> None:
    """
    Renderiza el análisis R.B. con meses en filas y tiendas en columnas.
    En cada mes:
      - R.B. más alto = verde
      - R.B. más bajo = rojo
    La columna TOTAL GRUPO no participa en la comparación.
    """
    if df_display.empty:
        st.info("No hay datos para mostrar con los filtros seleccionados.")
        return

    # Diseño compacto para evitar desplazamientos verticales innecesarios.
    altura_fila = min(int(altura_fila), 27)
    altura_cabecera = min(int(altura_cabecera), 30)

    verdes_por_columna = {t: set() for t in tiendas if t in df_display.columns}
    rojos_por_columna = {t: set() for t in tiendas if t in df_display.columns}

    for idx, fila in df_numericos.iterrows():
        etiqueta = str(fila["Mes"])

        # Totales/promedios/acumulados no participan en máximos/mínimos.
        etiqueta_norm = etiqueta.strip().upper()
        if (
            "TOTAL" in etiqueta_norm
            or "PROMEDIO" in etiqueta_norm
            or "ACUMULADO" in etiqueta_norm
        ):
            continue

        valores = {}

        for tienda in tiendas:
            if tienda not in df_numericos.columns:
                continue
            try:
                valor = float(fila[tienda])
                if pd.notna(valor):
                    valores[tienda] = valor
            except (TypeError, ValueError):
                pass

        if len(valores) < 2:
            continue

        minimo = min(valores.values())
        maximo = max(valores.values())

        if minimo == maximo:
            continue

        for tienda, valor in valores.items():
            if valor == maximo:
                verdes_por_columna[tienda].add(etiqueta)
            if valor == minimo:
                rojos_por_columna[tienda].add(etiqueta)

    gb = GridOptionsBuilder.from_dataframe(df_display)
    # El usuario puede redimensionar arrastrando el borde de la cabecera
    # y reordenar columnas arrastrando la propia cabecera.
    gb.configure_default_column(
        resizable=True,
        sortable=False,
        filter=False,
        suppressMovable=False,
    )
    gb.configure_default_column(
        resizable=True,
        filterable=False,
        sortable=False,
        editable=False,
        suppressMenu=True,
        wrapText=False,
        autoHeight=False,
    )

    for i, col in enumerate(df_display.columns):
        if col == "Mes":
            ancho = calcular_ancho_columna(df_display, col, 95)
            ancho_primera = max(ancho, 105)
            gb.configure_column(
                col,
                width=ancho_primera,
                minWidth=100,
                maxWidth=145,
                cellStyle={"textAlign": "left", "fontWeight": "600"},
            )
            continue

        ancho = calcular_ancho_columna(df_display, col, 78)

        col_norm = str(col).strip().upper()
        es_columna_total = es_columna_resumen(col)

        if es_columna_total:
            estilo = JsCode(
                r"""
                function(params) {
                    const texto = params.value === null || params.value === undefined
                        ? ''
                        : String(params.value).trim();

                    let estilo = {
                        'textAlign': 'right',
                        'backgroundColor': '#fff7cc',
                        'fontWeight': '700'
                    };

                    if (texto.startsWith('-') || /^\(.*\)$/.test(texto)) {
                        estilo['color'] = '#d00000';
                    }

                    return estilo;
                }
                """
            )

        elif col in tiendas:
            verdes = ",".join(repr(x) for x in sorted(verdes_por_columna.get(col, set())))
            rojos = ",".join(repr(x) for x in sorted(rojos_por_columna.get(col, set())))

            estilo = JsCode(
                f"""
                function(params) {{
                    const mes = params.data && params.data.Mes ? String(params.data.Mes) : '';
                    const verdes = [{verdes}];
                    const rojos = [{rojos}];

                    let estilo = {{'textAlign': 'right'}};

                    if (verdes.includes(mes)) {{
                        estilo['backgroundColor'] = '#d9ead3';
                        estilo['fontWeight'] = '700';
                    }} else if (rojos.includes(mes)) {{
                        estilo['backgroundColor'] = '#f4cccc';
                        estilo['fontWeight'] = '700';
                    }}

                    const texto = params.value === null || params.value === undefined
                        ? ''
                        : String(params.value).trim();

                    if (texto.startsWith('-') || /^\\(.*\\)$/.test(texto)) {{
                        estilo['color'] = '#d00000';
                    }}

                    return estilo;
                }}
                """
            )
        else:
            estilo = JsCode(
                r"""
                function(params) {
                    const texto = params.value === null || params.value === undefined
                        ? ''
                        : String(params.value).trim();

                    let estilo = {'textAlign': 'right'};

                    if (texto.startsWith('-') || /^\(.*\)$/.test(texto)) {
                        estilo['color'] = '#d00000';
                    }

                    return estilo;
                }
                """
            )

        gb.configure_column(
            col,
            width=ancho,
            minWidth=57,
            maxWidth=138,
            cellStyle=estilo,
        )

    # Sombrear la fila ACUMULADO
    get_row_style = JsCode(
        """
        function(params) {
            const mes = params.data && params.data.Mes ? String(params.data.Mes) : '';
            if (mes === 'ACUMULADO') {
                return {
                    'fontWeight': '700',
                    'backgroundColor': '#e9ecef'
                };
            }
            return null;
        }
        """
    )

    gb.configure_grid_options(
        domLayout="normal",
        suppressRowClickSelection=True,
        rowHeight=altura_fila,
        headerHeight=altura_cabecera,
        getRowStyle=get_row_style,
        suppressHorizontalScroll=False,
    )

    altura_tabla = altura_cabecera + (len(df_display) * altura_fila)

    firma_columnas = "__".join(str(c) for c in df_display.columns)

    # La preferencia de columnas pertenece a ESTA consulta concreta.
    # Así, al cambiar año o meses, la nueva información arranca siempre
    # con autoajuste inicial y no hereda anchos de una consulta anterior.
    ano_actual = globals().get("ano", "")
    meses_actuales = globals().get("meses_sel", [])
    firma_consulta = f"{ano_actual}__{'__'.join(str(m) for m in meses_actuales)}"

    clave_preferencias_efectiva = (
        f"{clave_preferencias}__{firma_consulta}__{firma_columnas}"
    )

    boton_autoajuste = st.button(
        "↔ Autoajustar todas las columnas",
        key=f"autoajustar_todas_{clave_preferencias_efectiva}",
        help="Ajusta automáticamente todas las columnas según su cabecera y contenido.",
    )

    token_key = f"token_autoajuste_{clave_preferencias_efectiva}"
    if token_key not in st.session_state:
        st.session_state[token_key] = 0

    if boton_autoajuste:
        st.session_state[token_key] += 1

    autoajustar_todas = boton_autoajuste

    gb.configure_grid_options(
        onGridReady=estado_columnas_js(
            clave_preferencias_efectiva,
            autoajustar_todas=autoajustar_todas,
        )
    )

    AgGrid(
        df_display,
        gridOptions=gb.build(),
        key=f"grid_{clave_preferencias_efectiva}_{st.session_state[token_key]}",
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True,
        theme="balham",
        height=altura_tabla,
        custom_css={
            ".ag-cell": {
                "font-size": "13px",
                "line-height": "31px",
                "padding-left": "3px",
                "padding-right": "3px",
                "border-right": "1px solid #c9ced3",
            },
            ".ag-header-cell": {
                "font-size": "13px",
                "font-weight": "600",
                "padding-left": "3px",
                "padding-right": "3px",
                "border-right": "1px solid #b8bec5",
            },
            ".ag-header-cell-label": {
                "justify-content": "center !important",
                "text-align": "center !important",
                "width": "100%",
            },
            ".ag-header-cell-text": {
                "text-align": "center !important",
                "width": "100%",
            },
        },
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
    df = load_data(firma_archivo_datos())
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
        "Análisis de Inventario y Rotación",
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

# Departamentos disponibles.
# Regla global: en cualquier pantalla se excluyen las tiendas que no tengan
# ningún movimiento en el año y meses seleccionados.
df_ano = df[df["Año"] == ano] if "Año" in df.columns else df

if "Mes" in df_ano.columns and meses_sel:
    df_periodo_departamentos = df_ano[df_ano["Mes"].isin(meses_sel)].copy()
else:
    df_periodo_departamentos = df_ano.copy()

departamentos_disponibles = []
if "Departamento" in df_periodo_departamentos.columns:
    for departamento in sorted(
        df_periodo_departamentos["Departamento"].dropna().astype(str).str.strip().unique()
    ):
        df_dep = df_periodo_departamentos[
            df_periodo_departamentos["Departamento"].astype(str).str.strip() == departamento
        ]

        if "Importe D" in df_dep.columns:
            movimiento = pd.to_numeric(
                df_dep["Importe D"], errors="coerce"
            ).fillna(0).abs().sum()
            tiene_datos = movimiento > 1e-12
        else:
            columnas_numericas = [
                c for c in df_dep.columns
                if pd.api.types.is_numeric_dtype(df_dep[c])
                and c != "Año"
            ]
            tiene_datos = bool(
                columnas_numericas
                and df_dep[columnas_numericas].fillna(0).abs().to_numpy().sum() > 1e-12
            )

        if tiene_datos:
            departamentos_disponibles.append(departamento)

departamentos_sin_general = [
    d for d in departamentos_disponibles
    if str(d).strip().upper() != "GENERAL"
]

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

    orientacion_rb = st.sidebar.radio(
        "Orientación de la tabla R.B.",
        [
            "Meses en filas / Tiendas en columnas",
            "Tiendas en filas / Meses en columnas",
        ],
    )
    
    # Para el análisis R.B. solo se muestran tiendas que tengan algún
    # valor de R.B. distinto de cero en los meses seleccionados.
    tiendas_rb_disponibles = []
    for tienda in departamentos_sin_general:
        tiene_rb = False
        for mes in meses_sel:
            df_rb_chk = obtener_filtro_datos(df, ano, [mes], [tienda])
            valor_rb_chk = calcular_rb_puro(df_rb_chk)
            if abs(float(valor_rb_chk)) > 1e-12:
                tiene_rb = True
                break

        if tiene_rb:
            tiendas_rb_disponibles.append(tienda)

    tiendas_rb = st.sidebar.multiselect(
        "Selecciona tiendas",
        tiendas_rb_disponibles,
        default=tiendas_rb_disponibles,
    )
    
    if not tiendas_rb or not meses_sel:
        st.warning("Selecciona al menos una tienda y un mes.")
        st.stop()
    
    # =====================================================================
    # EVOLUCIÓN MENSUAL POR TIENDA
    # =====================================================================
    if tipo_analisis_rb == "Evolución Mensual por Tienda":
        st.subheader(f"Análisis R.B. - Evolución Mensual por Tienda ({ano})")

        if orientacion_rb == "Meses en filas / Tiendas en columnas":
            # =============================================================
            # ORIENTACIÓN 1: Meses en filas / Tiendas en columnas
            # =============================================================
            columnas_tabla = ["Mes"] + tiendas_rb
            if len(tiendas_rb) > 1:
                columnas_tabla.append("TOTAL GRUPO")

            filas_display = []
            filas_nums = []

            for mes in meses_sel:
                fila_d = [mes]
                fila_n = [mes]

                for tienda in tiendas_rb:
                    df_filtrado = obtener_filtro_datos(df, ano, [mes], [tienda])
                    val_rb = calcular_rb_puro(df_filtrado)
                    fila_n.append(val_rb)
                    fila_d.append(formato_porcentaje(val_rb))

                if len(tiendas_rb) > 1:
                    df_total = obtener_filtro_datos(df, ano, [mes], tiendas_rb)
                    val_total = calcular_rb_puro(df_total)
                    fila_n.append(val_total)
                    fila_d.append(formato_porcentaje(val_total))

                filas_display.append(fila_d)
                filas_nums.append(fila_n)

            if len(meses_sel) > 1:
                fila_d = ["ACUMULADO"]
                fila_n = ["ACUMULADO"]

                for tienda in tiendas_rb:
                    df_acum = obtener_filtro_datos(df, ano, meses_sel, [tienda])
                    val_acum = calcular_rb_puro(df_acum)
                    fila_n.append(val_acum)
                    fila_d.append(formato_porcentaje(val_acum))

                if len(tiendas_rb) > 1:
                    df_acum_total = obtener_filtro_datos(df, ano, meses_sel, tiendas_rb)
                    val_acum_total = calcular_rb_puro(df_acum_total)
                    fila_n.append(val_acum_total)
                    fila_d.append(formato_porcentaje(val_acum_total))

                filas_display.append(fila_d)
                filas_nums.append(fila_n)

            df_res_d = pd.DataFrame(filas_display, columns=columnas_tabla)
            df_res_n = pd.DataFrame(filas_nums, columns=columnas_tabla)

            render_aggrid_rb_horizontal(
                df_res_d,
                df_res_n,
                tiendas_rb,
                altura_fila=36,
                altura_cabecera=38,
                clave_preferencias="rb_meses_filas",
            )

        else:
            # =============================================================
            # ORIENTACIÓN 2: Tiendas en filas / Meses en columnas
            # =============================================================
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

            render_aggrid_table(
                df_res_d,
                modo="auto",
                df_numericos=df_res_n,
                resaltar_rb_tiendas=len(tiendas_rb) > 1,
                columnas_comparar=tiendas_rb if len(tiendas_rb) > 1 else None,
                altura_fila=36,
                altura_cabecera=38,
        clave_preferencias="rb_tiendas_filas",
            )

        descargar_excel(
            df_res_n,
            "Analisis_RB_Mensual",
            f"Analisis_RB_Mensual_{ano}.xlsx",
            "Descargar Análisis R.B. en Excel",
        )

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
        
        render_aggrid_table(
            df_acum_d,
            modo="auto",
            df_numericos=df_acum_n,
            resaltar_rb_tiendas=len(tiendas_rb) > 1,
            columnas_comparar=tiendas_rb if len(tiendas_rb) > 1 else None,
            altura_fila=36,
            altura_cabecera=38,
        clave_preferencias="rb_tiendas_filas",
        )
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
        
        render_aggrid_table(
            df_inter_d,
            modo="auto",
            df_numericos=df_inter_n,
            resaltar_rb_tiendas=len(tiendas_rb) > 1,
            columnas_comparar=tiendas_rb if len(tiendas_rb) > 1 else None,
            altura_fila=36,
            altura_cabecera=38,
        clave_preferencias="rb_tiendas_filas",
        )
        descargar_excel(df_inter_n, "Interanual_RB", f"Comparativa_Interanual_RB_{ano}.xlsx",
                       "Descargar Comparativa R.B. en Excel")

# =====================================================================
# MÓDULO 2: INFORME KPI (% SOBRE VENTAS)
# =====================================================================

elif modulo_principal == "Informe KPI (% sobre Ventas)":
    base_kpi = st.sidebar.radio(
        "Base del KPI",
        ["% sobre Ventas", "€/m² de tienda", "Comparar ambos"],
    )
    m2_por_tienda = load_tiendas_m2(firma_archivo_datos())

    modo_analisis = st.sidebar.radio(
        "Tipo de Análisis KPI",
        [
            "Evolución Mensual / Tienda",
            "Comparativa Multi-Tienda (Totales)",
            "Comparativa Interanual (Año vs Año Anterior)",
        ],
    )
    
    # Selección de tiendas/meses.
    # Para KPI por m² la lista se construye DESDE EL PRINCIPIO solo con
    # tiendas válidas: nunca GENERAL, nunca tiendas sin movimiento y siempre
    # con m² informado en la hoja Tiendas.
    def tienda_kpi_tiene_datos(tienda: str) -> bool:
        """
        Una tienda solo se considera válida para KPI si, en el periodo
        seleccionado, alguno de los conceptos KPI calculados es distinto de cero.
        Esto evita incluir tiendas con movimientos residuales que no generan KPI.
        """
        df_tienda = obtener_filtro_datos(df, ano, meses_sel, [tienda])
        if df_tienda.empty:
            return False

        resultados_tienda = calcular_resultados(df_tienda)

        for concepto, valor in resultados_tienda.items():
            try:
                if abs(float(valor)) > 1e-12:
                    return True
            except (TypeError, ValueError):
                continue

        return False

    tiendas_kpi_normales = [
        t for t in departamentos_disponibles
        if str(t).strip().upper() != "GENERAL"
        and tienda_kpi_tiene_datos(t)
    ]

    tiendas_kpi_m2 = [
        t for t in tiendas_kpi_normales
        if m2_por_tienda.get(str(t).strip(), 0) > 0
    ]

    if base_kpi == "% sobre Ventas":
        opciones_tiendas_kpi = departamentos_disponibles
        default_tiendas_kpi = tiendas_kpi_normales
    else:
        opciones_tiendas_kpi = tiendas_kpi_m2
        default_tiendas_kpi = tiendas_kpi_m2

        if not m2_por_tienda:
            st.error(
                "No se ha podido leer la hoja 'Tiendas' con los metros cuadrados. "
                f"Archivo detectado: {obtener_archivo_datos()}"
            )
            st.stop()

        if not tiendas_kpi_m2:
            st.error(
                "No hay tiendas válidas para el KPI por m² en el año/mes seleccionado. "
                "Comprueba que tengan movimiento y m² informado en la hoja 'Tiendas'."
            )
            st.stop()

    if modo_analisis == "Comparativa Multi-Tienda (Totales)":
        tiendas = st.sidebar.multiselect(
            "Selecciona tiendas a comparar",
            opciones_tiendas_kpi,
            default=(
                default_tiendas_kpi
                if base_kpi != "% sobre Ventas"
                else default_tiendas_kpi[:2]
            ),
        )
    else:
        tipo_consulta = st.sidebar.radio(
            "Tipo de consulta", ["Una tienda", "Conjunto de tiendas"]
        )

        if tipo_consulta == "Una tienda":
            tienda_sel = st.sidebar.selectbox(
                "Selecciona tienda",
                opciones_tiendas_kpi,
            )
            tiendas = [tienda_sel] if tienda_sel else []
        else:
            tiendas = st.sidebar.multiselect(
                "Selecciona tiendas",
                opciones_tiendas_kpi,
                default=default_tiendas_kpi,
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
        
        render_aggrid_table(
            df_kpi_display,
            modo="auto",
            df_numericos=df_kpi_numericos,
            resaltar_kpi_tiendas=len([
                c for c in df_kpi_numericos.columns
                if c != "Resultados"
                and not es_columna_resumen(c)
            ]) > 1,
            columnas_comparar=[
                c for c in df_kpi_numericos.columns
                if c != "Resultados"
                and not es_columna_resumen(c)
            ],
            clave_preferencias="informe_kpi",
        )
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
        
        # Transformación opcional del KPI a €/m².
        if base_kpi != "% sobre Ventas":
            if modo_analisis == "Comparativa Multi-Tienda (Totales)" or (
                modo_analisis == "Evolución Mensual / Tienda" and len(tiendas) > 1
            ):
                columnas_originales = [
                    c for c in df_kpi_numericos.columns
                    if c != "Resultados" and not es_columna_resumen(c)
                ]
                for c in columnas_originales:
                    metros = m2_por_tienda.get(str(c).strip(), 0)
                    if metros > 0 and c in datos_fuente:
                        res_c = datos_fuente[c]
                        for idx, concepto in enumerate(df_kpi_numericos["Resultados"]):
                            if concepto != "R. B.":
                                valor_m2 = res_c.get(concepto, 0.0) / metros
                                if base_kpi == "€/m² de tienda":
                                    df_kpi_numericos.at[idx, c] = valor_m2
                                    df_kpi_display.at[idx, c] = formato_moneda(valor_m2)

                if base_kpi == "Comparar ambos":
                    disp = df_kpi_display[["Resultados"]].copy()
                    nums = df_kpi_numericos[["Resultados"]].copy()
                    # Recuperar porcentajes desde datos_fuente y añadir ambas medidas.
                    for c in columnas_originales:
                        if c not in datos_fuente:
                            continue
                        res_c = datos_fuente[c]
                        ventas_c = res_c.get("Ventas", 0.0)
                        metros = m2_por_tienda.get(str(c).strip(), 0)
                        pct_vals, pct_disp, m2_vals, m2_disp = [], [], [], []
                        for concepto in df_kpi_numericos["Resultados"]:
                            valor = res_c.get(concepto, 0.0)
                            pct = valor if concepto == "R. B." else (
                                valor / ventas_c if ventas_c else 0.0
                            )
                            vm2 = 0.0 if concepto == "R. B." else (
                                valor / metros if metros else 0.0
                            )
                            pct_vals.append(pct); pct_disp.append(formato_porcentaje(pct))
                            m2_vals.append(vm2)
                            m2_disp.append(formato_porcentaje(pct) if concepto == "R. B." else formato_moneda(vm2))
                        nums[f"{c} %"] = pct_vals
                        disp[f"{c} %"] = pct_disp
                        nums[f"{c} €/m²"] = m2_vals
                        disp[f"{c} €/m²"] = m2_disp
                    df_kpi_display, df_kpi_numericos = disp, nums

        render_aggrid_table(
            df_kpi_display,
            modo="auto",
            df_numericos=df_kpi_numericos,
            resaltar_kpi_tiendas=len([
                c for c in df_kpi_numericos.columns
                if c != "Resultados"
                and not es_columna_resumen(c)
            ]) > 1,
            columnas_comparar=[
                c for c in df_kpi_numericos.columns
                if c != "Resultados"
                and not es_columna_resumen(c)
            ],
            clave_preferencias="informe_kpi",
        )
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
        
        # Transformación opcional del KPI a €/m².
        if base_kpi != "% sobre Ventas":
            if modo_analisis == "Comparativa Multi-Tienda (Totales)" or (
                modo_analisis == "Evolución Mensual / Tienda" and len(tiendas) > 1
            ):
                columnas_originales = [
                    c for c in df_kpi_numericos.columns
                    if c != "Resultados" and not es_columna_resumen(c)
                ]
                for c in columnas_originales:
                    metros = m2_por_tienda.get(str(c).strip(), 0)
                    if metros > 0 and c in datos_fuente:
                        res_c = datos_fuente[c]
                        for idx, concepto in enumerate(df_kpi_numericos["Resultados"]):
                            if concepto != "R. B.":
                                valor_m2 = res_c.get(concepto, 0.0) / metros
                                if base_kpi == "€/m² de tienda":
                                    df_kpi_numericos.at[idx, c] = valor_m2
                                    df_kpi_display.at[idx, c] = formato_moneda(valor_m2)

                if base_kpi == "Comparar ambos":
                    disp = df_kpi_display[["Resultados"]].copy()
                    nums = df_kpi_numericos[["Resultados"]].copy()
                    # Recuperar porcentajes desde datos_fuente y añadir ambas medidas.
                    for c in columnas_originales:
                        if c not in datos_fuente:
                            continue
                        res_c = datos_fuente[c]
                        ventas_c = res_c.get("Ventas", 0.0)
                        metros = m2_por_tienda.get(str(c).strip(), 0)
                        pct_vals, pct_disp, m2_vals, m2_disp = [], [], [], []
                        for concepto in df_kpi_numericos["Resultados"]:
                            valor = res_c.get(concepto, 0.0)
                            pct = valor if concepto == "R. B." else (
                                valor / ventas_c if ventas_c else 0.0
                            )
                            vm2 = 0.0 if concepto == "R. B." else (
                                valor / metros if metros else 0.0
                            )
                            pct_vals.append(pct); pct_disp.append(formato_porcentaje(pct))
                            m2_vals.append(vm2)
                            m2_disp.append(formato_porcentaje(pct) if concepto == "R. B." else formato_moneda(vm2))
                        nums[f"{c} %"] = pct_vals
                        disp[f"{c} %"] = pct_disp
                        nums[f"{c} €/m²"] = m2_vals
                        disp[f"{c} €/m²"] = m2_disp
                    df_kpi_display, df_kpi_numericos = disp, nums

        render_aggrid_table(
            df_kpi_display,
            modo="auto",
            df_numericos=df_kpi_numericos,
            resaltar_kpi_tiendas=len([
                c for c in df_kpi_numericos.columns
                if c != "Resultados"
                and not es_columna_resumen(c)
            ]) > 1,
            columnas_comparar=[
                c for c in df_kpi_numericos.columns
                if c != "Resultados"
                and not es_columna_resumen(c)
            ],
            clave_preferencias="informe_kpi",
        )
        descargar_excel(df_kpi_numericos, "Informe_KPI", f"Informe_KPI_Ventas_{ano}.xlsx",
                       "Descargar Informe KPI en Excel")

# =====================================================================
# MÓDULO 3: CUENTA DE RESULTADOS COMPLETA
# =====================================================================

elif modulo_principal == "Cuenta de Resultados Completa":
    modo_analisis = st.sidebar.radio(
        "Tipo de Análisis",
        [
            "Evolución Mensual / Tienda",
            "Comparativa Multi-Tienda (Totales)",
            "Comparativa Interanual (Año vs Año Anterior)",
        ],
    )
    
    # Selección de tiendas
    total_empresa_consulta = False
    if modo_analisis == "Comparativa Multi-Tienda (Totales)":
        tiendas = st.sidebar.multiselect(
            "Selecciona tiendas a comparar",
            departamentos_sin_general,
            default=departamentos_sin_general[:2]
            if len(departamentos_sin_general) >= 2
            else departamentos_sin_general,
        )
    else:
        tipo_consulta = st.sidebar.radio(
            "Tipo de consulta", ["Total empresa", "Una tienda", "Conjunto de tiendas"]
        )
        if tipo_consulta == "Total empresa":
            total_empresa_consulta = True
            tiendas = departamentos_disponibles.copy()
        elif tipo_consulta == "Una tienda":
            tienda_sel = st.sidebar.selectbox("Selecciona tienda", departamentos_disponibles)
            tiendas = [tienda_sel] if tienda_sel else []
        else:
            tiendas = st.sidebar.multiselect(
                "Selecciona tiendas",
                departamentos_sin_general,
                default=departamentos_sin_general,
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
        
        datos_fuente["Ant"] = calcular_resultados_seleccion(
            df_ant, ano_anterior, meses_sel, tiendas, departamentos_sin_general
        )
        datos_fuente["Act"] = calcular_resultados_seleccion(
            df_act, ano, meses_sel, tiendas, departamentos_sin_general
        )
        
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
        datos_fuente["Total"] = calcular_resultados_seleccion(
            df_filtrado, ano, meses_sel, tiendas, departamentos_sin_general
        )
        
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
        
        if total_empresa_consulta:
            # TOTAL EMPRESA: meses en columnas, incluyendo General y ajustes mensuales.
            for mes in meses_sel:
                datos_fuente[mes] = calcular_resultados_total_empresa(
                    df, ano, [mes]
                )

            if len(meses_sel) > 1:
                datos_fuente["Total"] = calcular_resultados_total_empresa(
                    df, ano, meses_sel
                )
                columnas_eje = meses_sel + ["Total"]
            else:
                columnas_eje = meses_sel

        elif len(tiendas) > 1:
            # Comparativa por tiendas (sin repartir los ajustes globales).
            for tienda in tiendas:
                df_filtrado = obtener_filtro_datos(df, ano, meses_sel, [tienda])
                datos_fuente[tienda] = calcular_resultados(df_filtrado)

            df_filtrado = obtener_filtro_datos(df, ano, meses_sel, tiendas)
            datos_fuente["Total"] = calcular_resultados(df_filtrado)
            columnas_eje = tiendas + ["Total"]
        else:
            # Evolución mensual de una sola tienda.
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


# =====================================================================
# MÓDULO 4: ANÁLISIS DE INVENTARIO Y ROTACIÓN
# =====================================================================

elif modulo_principal == "Análisis de Inventario y Rotación":
    st.header("Análisis de Inventario y Rotación")

    # Los m² se cargan también en este módulo para poder calcular
    # Stock €/m² en la comparativa de tiendas.
    m2_por_tienda = load_tiendas_m2(firma_archivo_datos())
    df_inventario = load_inventario(firma_archivo_datos())

    if df_inventario.empty:
        st.error(
            "No se ha podido leer la hoja 'Inventario'. "
            f"Archivo detectado: {obtener_archivo_datos()}"
        )
        st.stop()

    vista_inventario = st.sidebar.radio(
        "Vista de inventario",
        [
            "Rotación / Cobertura",
            "Inventarios Mensuales",
            "Comparativa de Tiendas",
            "Eficiencia de Inventario",
        ],
        key="vista_inventario",
    )

    # -------------------------------------------------------------
    # VISTA: INVENTARIOS MENSUALES
    # -------------------------------------------------------------
    if vista_inventario == "Inventarios Mensuales":
        inv_periodo_mensual = df_inventario[
            (df_inventario["Año"] == ano)
            & (df_inventario["Mes"].isin(meses_sel))
        ].copy()

        if inv_periodo_mensual.empty:
            st.info(
                "No hay inventario informado para el año y los meses seleccionados."
            )
            st.stop()

        tiendas_mensuales = sorted(
            inv_periodo_mensual["Departamento"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        tiendas_sel_mensual = st.sidebar.multiselect(
            "Selecciona tiendas",
            tiendas_mensuales,
            default=tiendas_mensuales,
            key="tiendas_inventario_mensual",
        )

        if not tiendas_sel_mensual:
            st.info("Selecciona al menos una tienda.")
            st.stop()

        inv_sel = inv_periodo_mensual[
            inv_periodo_mensual["Departamento"]
            .astype(str)
            .str.strip()
            .isin(tiendas_sel_mensual)
        ].copy()

        tabla_inv = (
            inv_sel.pivot_table(
                index="Departamento",
                columns="Mes",
                values="Inventario",
                aggfunc="sum",
                fill_value=0,
            )
            .reindex(columns=[m for m in MESES_ORDEN if m in meses_sel])
        )

        # TOTAL por mes, calculado con las tiendas seleccionadas.
        tabla_inv.loc["TOTAL"] = tabla_inv.sum(axis=0)

        tabla_inv = tabla_inv.reset_index().rename(
            columns={"Departamento": "Tienda"}
        )

        tabla_inv_numericos = tabla_inv.copy()
        tabla_inv_display = tabla_inv.copy()

        for col in tabla_inv_display.columns:
            if col != "Tienda":
                tabla_inv_display[col] = tabla_inv_display[col].apply(
                    formato_moneda
                )

        st.subheader(f"Inventarios Mensuales — {ano}")
        st.caption(
            "Inventario final de cada mes por tienda. "
            "La fila TOTAL se recalcula con las tiendas seleccionadas."
        )

        render_aggrid_table(
            tabla_inv_display,
            modo="auto",
            altura_fila=32,
            altura_cabecera=38,
            clave_preferencias="inventarios_mensuales",
        )

        descargar_excel(
            tabla_inv_numericos,
            "Inventarios Mensuales",
            f"Inventarios_Mensuales_{ano}.xlsx",
            "Descargar inventarios mensuales en Excel",
        )

        st.stop()

    # -------------------------------------------------------------
    # VISTA: COMPARATIVA DE TIENDAS
    # -------------------------------------------------------------
    if vista_inventario == "Comparativa de Tiendas":
        inv_periodo_cmp = df_inventario[
            (df_inventario["Año"] == ano)
            & (df_inventario["Mes"].isin(meses_sel))
        ].copy()

        if inv_periodo_cmp.empty:
            st.info("No hay inventario informado para el periodo seleccionado.")
            st.stop()

        # GENERAL está disponible para seleccionarlo, pero no aparece
        # seleccionado por defecto en la comparativa.
        tiendas_cmp = sorted(
            inv_periodo_cmp["Departamento"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )
        tiendas_cmp_default = [
            tienda for tienda in tiendas_cmp
            if str(tienda).strip().upper() != "GENERAL"
        ]

        tiendas_cmp_sel = st.sidebar.multiselect(
            "Selecciona tiendas",
            tiendas_cmp,
            default=tiendas_cmp_default,
            key="tiendas_cmp_inventario",
        )

        if not tiendas_cmp_sel:
            st.info("Selecciona al menos una tienda.")
            st.stop()

        filas_cmp = []
        for tienda in tiendas_cmp_sel:
            df_inv_tienda = inv_periodo_cmp[
                inv_periodo_cmp["Departamento"].astype(str).str.strip()
                == str(tienda).strip()
            ]
            inventario_medio = float(
                pd.to_numeric(
                    df_inv_tienda["Inventario"], errors="coerce"
                ).dropna().mean()
            ) if not df_inv_tienda.empty else 0.0

            inventario_ultimo = 0.0
            if not df_inv_tienda.empty:
                orden_mes = {m: i for i, m in enumerate(MESES_ORDEN)}
                aux = df_inv_tienda.copy()
                aux["_orden"] = aux["Mes"].map(orden_mes)
                aux = aux.sort_values("_orden")
                inventario_ultimo = float(
                    pd.to_numeric(
                        aux["Inventario"], errors="coerce"
                    ).fillna(0).iloc[-1]
                )

            df_ventas = obtener_filtro_datos(
                df, ano, meses_sel, [tienda]
            )
            resultados_tienda = calcular_resultados(df_ventas)
            ventas = float(resultados_tienda.get("Ventas", 0.0))
            margen_bruto = float(resultados_tienda.get("MARGEN BRUTO", 0.0))
            coste_ventas = float(resultados_tienda.get("Coste Ventas", 0.0))

            metros = m2_por_tienda.get(str(tienda).strip(), 0)
            stock_m2 = inventario_ultimo / metros if metros > 0 else None
            ventas_stock = ventas / inventario_medio if abs(inventario_medio) > 1e-12 else None
            margen_stock = margen_bruto / inventario_medio if abs(inventario_medio) > 1e-12 else None
            rotacion_coste_stock = (
                abs(coste_ventas) / inventario_medio
                if abs(inventario_medio) > 1e-12 else None
            )

            filas_cmp.append(
                {
                    "Tienda": tienda,
                    "Inventario Último": inventario_ultimo,
                    "Inventario Medio": inventario_medio,
                    "Ventas Periodo": ventas,
                    "Margen Bruto": margen_bruto,
                    "Stock €/m²": stock_m2,
                    "Ventas / Stock": ventas_stock,
                    "Margen / Stock": margen_stock,
                    "Rotación Coste / Stock": rotacion_coste_stock,
                }
            )

        df_cmp_num = pd.DataFrame(filas_cmp)

        # ---------------------------------------------------------
        # TOTAL de las tiendas seleccionadas
        # ---------------------------------------------------------
        # Inventario último: suma del último inventario de cada tienda.
        inventario_ultimo_total = float(
            pd.to_numeric(df_cmp_num["Inventario Último"], errors="coerce")
            .fillna(0)
            .sum()
        )

        # Inventario medio TOTAL: primero se suma el inventario por mes
        # de las tiendas seleccionadas y después se calcula la media mensual.
        inv_sel_total = inv_periodo_cmp[
            inv_periodo_cmp["Departamento"].astype(str).str.strip().isin(
                [str(x).strip() for x in tiendas_cmp_sel]
            )
        ].copy()

        inv_mes_total = (
            inv_sel_total.groupby("Mes", as_index=False)["Inventario"].sum()
            if not inv_sel_total.empty
            else pd.DataFrame(columns=["Mes", "Inventario"])
        )

        inventario_medio_total = float(
            pd.to_numeric(inv_mes_total["Inventario"], errors="coerce")
            .dropna()
            .mean()
        ) if not inv_mes_total.empty else 0.0

        # Ventas y margen bruto: suma de las tiendas seleccionadas.
        ventas_total = float(
            pd.to_numeric(df_cmp_num["Ventas Periodo"], errors="coerce")
            .fillna(0)
            .sum()
        )
        margen_total = float(
            pd.to_numeric(df_cmp_num["Margen Bruto"], errors="coerce")
            .fillna(0)
            .sum()
        )

        # Stock €/m² TOTAL:
        # solo se incluyen en este ratio las tiendas con m² informados.
        tiendas_con_m2 = [
            str(x).strip()
            for x in tiendas_cmp_sel
            if float(m2_por_tienda.get(str(x).strip(), 0) or 0) > 0
        ]
        metros_total = sum(
            float(m2_por_tienda.get(tienda, 0) or 0)
            for tienda in tiendas_con_m2
        )

        inventario_ultimo_con_m2 = 0.0
        for _, fila_cmp in df_cmp_num.iterrows():
            tienda_cmp = str(fila_cmp["Tienda"]).strip()
            if tienda_cmp in tiendas_con_m2:
                inventario_ultimo_con_m2 += float(
                    pd.to_numeric(fila_cmp["Inventario Último"], errors="coerce")
                    if pd.notna(pd.to_numeric(fila_cmp["Inventario Último"], errors="coerce"))
                    else 0.0
                )

        stock_m2_total = (
            inventario_ultimo_con_m2 / metros_total
            if metros_total > 0
            else None
        )

        ventas_stock_total = (
            ventas_total / inventario_medio_total
            if abs(inventario_medio_total) > 1e-12
            else None
        )
        margen_stock_total = (
            margen_total / inventario_medio_total
            if abs(inventario_medio_total) > 1e-12
            else None
        )

        coste_ventas_total = 0.0
        for tienda in tiendas_cmp_sel:
            df_ventas_tienda_total = obtener_filtro_datos(
                df, ano, meses_sel, [tienda]
            )
            resultados_tienda_total = calcular_resultados(df_ventas_tienda_total)
            coste_ventas_total += abs(
                float(resultados_tienda_total.get("Coste Ventas", 0.0))
            )

        rotacion_coste_stock_total = (
            coste_ventas_total / inventario_medio_total
            if abs(inventario_medio_total) > 1e-12
            else None
        )

        fila_total_cmp = pd.DataFrame([{
            "Tienda": "TOTAL",
            "Inventario Último": inventario_ultimo_total,
            "Inventario Medio": inventario_medio_total,
            "Ventas Periodo": ventas_total,
            "Margen Bruto": margen_total,
            "Stock €/m²": stock_m2_total,
            "Ventas / Stock": ventas_stock_total,
            "Margen / Stock": margen_stock_total,
            "Rotación Coste / Stock": rotacion_coste_stock_total,
        }])

        df_cmp_num = pd.concat(
            [df_cmp_num, fila_total_cmp],
            ignore_index=True,
        )

        df_cmp_disp = df_cmp_num.copy()

        for col in ["Inventario Último", "Inventario Medio", "Ventas Periodo", "Margen Bruto", "Stock €/m²"]:
            if col in df_cmp_disp.columns:
                df_cmp_disp[col] = df_cmp_disp[col].apply(
                    lambda x: formato_moneda(x) if pd.notna(x) else ""
                )

        for col in ["Ventas / Stock", "Margen / Stock", "Rotación Coste / Stock"]:
            if col in df_cmp_disp.columns:
                df_cmp_disp[col] = df_cmp_disp[col].apply(
                    lambda x: (
                        f"{float(x):,.2f}"
                        .replace(",", "X")
                        .replace(".", ",")
                        .replace("X", ".")
                    ) if pd.notna(x) else ""
                )

        st.subheader(f"Comparativa de Tiendas — {ano}")
        st.caption(
            "Inventario, ventas, margen y eficiencia del stock para el periodo seleccionado. "
            "La fila TOTAL recalcula los ratios sobre el conjunto de tiendas seleccionadas."
        )

        with st.expander("ℹ️ Qué significa cada dato", expanded=False):
            st.markdown(
                """
**Tienda**  
Establecimiento analizado. **General no aparece seleccionado por defecto**, pero puede añadirse manualmente desde el selector de tiendas. La fila **TOTAL** representa el conjunto de tiendas que estén seleccionadas.

**Inventario Último**  
Inventario final del último mes incluido en la selección.  
En **TOTAL** es la suma de los inventarios finales de las tiendas seleccionadas.

**Inventario Medio**  
Promedio del inventario final de los meses seleccionados.  
En **TOTAL** se suma primero el inventario de todas las tiendas en cada mes y después se calcula la media mensual.

**Ventas Periodo**  
Ventas acumuladas durante los meses seleccionados.  
En **TOTAL** es la suma de las ventas de las tiendas seleccionadas.

**Margen Bruto**  
Margen bruto generado durante el periodo seleccionado.  
En **TOTAL** es la suma del margen bruto de las tiendas seleccionadas.

**Stock €/m²**  
Inventario del último mes dividido entre los metros cuadrados de la tienda.  
Permite comparar cuánto stock mantiene cada establecimiento por unidad de superficie.  
En **TOTAL** se calcula únicamente con las tiendas que tienen m² informados; las tiendas sin superficie válida no entran en este ratio.

**Ventas / Stock**  
Ventas del periodo divididas entre el inventario medio.  
Ejemplo: **3,00** significa que por cada 1 € de inventario medio se han generado 3 € de ventas.  
En general, **mayor = mejor aprovechamiento comercial del stock**.

**Margen / Stock**  
Margen bruto del periodo dividido entre el inventario medio.  
Mide cuánto margen bruto genera cada euro mantenido de media en inventario.  
En general, **mayor = mejor eficiencia económica del stock**.

**Rotación Coste / Stock**  
Coste de ventas del periodo dividido entre el inventario medio.  
Mide cuántas veces rota el inventario a coste durante el periodo seleccionado.  
En general, **mayor = más rotación**, aunque un valor excesivamente alto puede indicar un stock demasiado ajustado.

### 🟢 Mejor / 🟠 Peor

En cada columna, el **mejor valor válido entre las tiendas seleccionadas aparece sombreado en verde** y el **peor en naranja**. Los valores **en blanco, sin dato o iguales a cero** quedan fuera de la comparación. La fila **TOTAL** tampoco participa en el semáforo.

En **Ventas Periodo, Margen Bruto, Ventas / Stock, Margen / Stock y Rotación Coste / Stock**, un valor mayor se considera mejor dentro de esta comparación. En las columnas de inventario y Stock €/m², el color identifica simplemente el **valor más alto y el más bajo**; no significa necesariamente que tener más o menos inventario sea bueno o malo por sí mismo.

**Cómo interpretar el TOTAL**  
Los ratios de la fila TOTAL **no se suman ni se promedian directamente**. Se vuelven a calcular utilizando los importes totales del conjunto de tiendas seleccionadas, para que el resultado sea coherente.
                """
            )

        render_aggrid_table(
            df_cmp_disp,
            modo="auto",
            df_numericos=df_cmp_num,
            resaltar_extremos_filas=True,
            columnas_extremos=[
                "Inventario Último",
                "Inventario Medio",
                "Ventas Periodo",
                "Margen Bruto",
                "Stock €/m²",
                "Ventas / Stock",
                "Margen / Stock",
                "Rotación Coste / Stock",
            ],
            altura_fila=32,
            altura_cabecera=38,
            clave_preferencias="comparativa_inventario",
        )

        # Gráfico inferior con indicador seleccionable
        indicadores_grafico_cmp = [
            "Ventas / Stock",
            "Margen / Stock",
            "Rotación Coste / Stock",
        ]

        indicador_grafico_cmp = st.selectbox(
            "Indicador del gráfico",
            indicadores_grafico_cmp,
            index=1,
            key="indicador_grafico_comparativa",
        )

        df_graf_cmp = df_cmp_num[
            (df_cmp_num["Tienda"].astype(str).str.strip().str.upper() != "TOTAL")
        ][["Tienda", indicador_grafico_cmp]].copy()

        df_graf_cmp[indicador_grafico_cmp] = pd.to_numeric(
            df_graf_cmp[indicador_grafico_cmp], errors="coerce"
        )
        df_graf_cmp = df_graf_cmp[
            df_graf_cmp[indicador_grafico_cmp].notna()
            & df_graf_cmp[indicador_grafico_cmp].ne(0)
        ].copy()

        if not df_graf_cmp.empty:
            df_graf_cmp = df_graf_cmp.sort_values(
                indicador_grafico_cmp, ascending=False
            )

            explicaciones_grafico_cmp = {
                "Ventas / Stock": (
                    "Ventas generadas durante el periodo por cada euro de inventario medio."
                ),
                "Margen / Stock": (
                    "Margen bruto generado durante el periodo por cada euro de inventario medio."
                ),
                "Rotación Coste / Stock": (
                    "Número de veces que el inventario rota a coste durante el periodo seleccionado."
                ),
            }

            st.markdown(f"#### {indicador_grafico_cmp} por tienda")
            st.caption(explicaciones_grafico_cmp[indicador_grafico_cmp])
            st.bar_chart(
                df_graf_cmp.set_index("Tienda")[indicador_grafico_cmp],
                use_container_width=True,
            )
        else:
            st.info(
                f"No hay datos válidos para representar {indicador_grafico_cmp} "
                "con la selección actual."
            )

        descargar_excel(
            df_cmp_num,
            "Comparativa Inventario",
            f"Comparativa_Inventario_{ano}.xlsx",
            "Descargar comparativa en Excel",
        )
        st.stop()

    # -------------------------------------------------------------
    # VISTA: EFICIENCIA DE INVENTARIO
    # -------------------------------------------------------------
    if vista_inventario == "Eficiencia de Inventario":
        inv_periodo_eff = df_inventario[
            (df_inventario["Año"] == ano)
            & (df_inventario["Mes"].isin(meses_sel))
        ].copy()

        if inv_periodo_eff.empty:
            st.info("No hay inventario informado para el periodo seleccionado.")
            st.stop()

        tiendas_eff = sorted(
            inv_periodo_eff["Departamento"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        tiendas_eff_sel = st.sidebar.multiselect(
            "Selecciona tiendas",
            tiendas_eff,
            default=tiendas_eff,
            key="tiendas_eff_inventario",
        )

        if not tiendas_eff_sel:
            st.info("Selecciona al menos una tienda.")
            st.stop()

        filas_eff = []

        for tienda in tiendas_eff_sel:
            if str(tienda).strip().upper() == "GENERAL":
                continue

            df_inv_tienda = inv_periodo_eff[
                inv_periodo_eff["Departamento"].astype(str).str.strip()
                == str(tienda).strip()
            ]

            inv_medio = float(
                pd.to_numeric(
                    df_inv_tienda["Inventario"], errors="coerce"
                ).dropna().mean()
            ) if not df_inv_tienda.empty else 0.0

            df_periodo = obtener_filtro_datos(df, ano, meses_sel, [tienda])
            res = calcular_resultados(df_periodo)

            ventas = float(res.get("Ventas", 0.0))
            coste = float(res.get("Coste Ventas", 0.0))
            margen = float(res.get("MARGEN BRUTO", res.get("R. B.", 0.0)))

            ventas_stock = ventas / inv_medio if abs(inv_medio) > 1e-12 else None
            margen_stock = margen / inv_medio if abs(inv_medio) > 1e-12 else None
            rotacion_stock = abs(coste) / inv_medio if abs(inv_medio) > 1e-12 else None

            filas_eff.append(
                {
                    "Tienda": tienda,
                    "Inventario Medio": inv_medio,
                    "Ventas Periodo": ventas,
                    "% Margen Acumulado": (margen / ventas) if abs(ventas) > 1e-12 else None,
                    "Margen Bruto Periodo": margen,
                    "Coste Ventas Periodo": coste,
                    "Ventas / Stock": ventas_stock,
                    "Margen / Stock": margen_stock,
                    "Rotación Coste / Stock": rotacion_stock,
                }
            )

        df_eff_num = pd.DataFrame(filas_eff)

        if df_eff_num.empty:
            st.info("No hay tiendas operativas para calcular eficiencia.")
            st.stop()

        # TOTAL EMPRESA:
        # Inventario medio = media de los inventarios totales mensuales,
        # incluyendo General. Ventas/coste se calculan sobre toda la empresa.
        totales_inv_mes = (
            inv_periodo_eff.groupby("Mes", as_index=False)["Inventario"].sum()
        )
        inv_medio_total = float(
            pd.to_numeric(totales_inv_mes["Inventario"], errors="coerce")
            .dropna()
            .mean()
        ) if not totales_inv_mes.empty else 0.0

        df_periodo_total = obtener_filtro_datos(df, ano, meses_sel, None)
        res_total = calcular_resultados(df_periodo_total)
        ventas_total = float(res_total.get("Ventas", 0.0))
        coste_total = float(res_total.get("Coste Ventas", 0.0))

        # Para el TOTAL, el margen acumulado oficial del ERP manda.
        # Se toma el margen acumulado correspondiente al último mes seleccionado.
        df_margenes_erp = load_margenes_totales_acumulados(firma_archivo_datos())
        meses_validos = [m for m in MESES_ORDEN if m in meses_sel]
        ultimo_mes = meses_validos[-1] if meses_validos else None
        margen_pct_total = None

        if ultimo_mes is not None and not df_margenes_erp.empty:
            fila_margen_erp = df_margenes_erp[
                (df_margenes_erp["Año"] == ano)
                & (df_margenes_erp["Mes"] == ultimo_mes)
            ]
            if not fila_margen_erp.empty:
                margen_pct_total = float(
                    fila_margen_erp["Margen Acumulado Total"].iloc[0]
                )
                if abs(margen_pct_total) > 1.0:
                    margen_pct_total = margen_pct_total / 100.0

        # Margen bruto TOTAL del periodo usando el margen acumulado ERP.
        # Así el total respeta exactamente el margen oficial acumulado.
        if margen_pct_total is not None:
            margen_total = ventas_total * margen_pct_total
            coste_total_erp = ventas_total - margen_total
        else:
            margen_total = float(res_total.get("MARGEN BRUTO", res_total.get("R. B.", 0.0)))
            coste_total_erp = coste_total

        ventas_stock_total = (
            ventas_total / inv_medio_total if abs(inv_medio_total) > 1e-12 else None
        )
        margen_stock_total = (
            margen_total / inv_medio_total if abs(inv_medio_total) > 1e-12 else None
        )
        rotacion_total = (
            abs(coste_total_erp) / inv_medio_total
            if abs(inv_medio_total) > 1e-12 else None
        )

        # Ranking solo para tiendas; TOTAL queda al final.
        df_eff_num["Ranking"] = (
            df_eff_num["Margen / Stock"]
            .rank(method="min", ascending=False)
        )
        df_eff_num = df_eff_num.sort_values(
            ["Ranking", "Tienda"]
        ).reset_index(drop=True)

        fila_total = pd.DataFrame([{
            "Tienda": "TOTAL",
            "Inventario Medio": inv_medio_total,
            "Ventas Periodo": ventas_total,
            "% Margen Acumulado": margen_pct_total,
            "Margen Bruto Periodo": margen_total,
            "Coste Ventas Periodo": coste_total_erp,
            "Ventas / Stock": ventas_stock_total,
            "Margen / Stock": margen_stock_total,
            "Rotación Coste / Stock": rotacion_total,
            "Ranking": None,
        }])

        df_eff_num = pd.concat([df_eff_num, fila_total], ignore_index=True)

        df_eff_disp = df_eff_num.copy()

        df_eff_disp["% Margen Acumulado"] = df_eff_disp["% Margen Acumulado"].apply(
            lambda x: formato_porcentaje(x) if pd.notna(x) else ""
        )

        for col in [
            "Inventario Medio",
            "Ventas Periodo",
            "Margen Bruto Periodo",
            "Coste Ventas Periodo",
        ]:
            df_eff_disp[col] = df_eff_disp[col].apply(
                lambda x: formato_moneda(x) if pd.notna(x) else ""
            )

        for col in ["Ventas / Stock", "Margen / Stock", "Rotación Coste / Stock"]:
            df_eff_disp[col] = df_eff_disp[col].apply(
                lambda x: (
                    f"{float(x):,.2f}"
                    .replace(",", "X")
                    .replace(".", ",")
                    .replace("X", ".")
                ) if pd.notna(x) else ""
            )

        df_eff_disp["Ranking"] = df_eff_disp["Ranking"].apply(
            lambda x: str(int(x)) if pd.notna(x) else ""
        )

        st.subheader(f"Eficiencia de Inventario — {ano}")
        st.caption(
            "Ranking de eficiencia según Margen Bruto / Inventario Medio. "
            "La fila TOTAL utiliza el margen acumulado oficial del ERP del último mes seleccionado."
        )

        with st.expander("ℹ️ Qué significa cada dato", expanded=False):
            st.markdown(
                """
**Tienda**  
Establecimiento analizado.

**Inventario Medio**  
Promedio del inventario final de los meses seleccionados. Indica cuánto dinero tiene inmovilizado de media la tienda en existencias.

**Ventas Periodo**  
Ventas acumuladas de la tienda durante los meses seleccionados.

**% Margen Acumulado**  
Porcentaje de margen bruto acumulado correspondiente al periodo seleccionado. Permite ver qué porcentaje de las ventas se convierte en margen bruto. En la fila **TOTAL** se muestra el porcentaje acumulado oficial del ERP correspondiente al último mes seleccionado.

**Margen Bruto Periodo**  
Margen bruto generado durante el periodo seleccionado. En la fila **TOTAL** se calcula utilizando el margen acumulado oficial del ERP del último mes seleccionado.

**Coste Ventas Periodo**  
Coste de la mercancía vendida durante el periodo. En la fila **TOTAL** se obtiene de forma coherente con el margen acumulado oficial del ERP.

**Ventas / Stock**  
Ventas del periodo divididas entre el inventario medio.  
Ejemplo: un valor de **3,00** significa que por cada 1 € de inventario medio se han generado 3 € de ventas.

**Margen / Stock**  
Margen bruto del periodo dividido entre el inventario medio.  
Es uno de los principales indicadores de eficiencia: cuanto **mayor** sea, mejor rendimiento económico se obtiene del stock.

**Rotación Coste / Stock**  
Coste de ventas dividido entre el inventario medio.  
Indica cuántas veces el coste de la mercancía vendida representa el inventario medio mantenido durante el periodo.

**Ranking**  
Ordena las tiendas según **Margen / Stock**.  
La posición **1** corresponde a la tienda que obtiene mayor margen bruto por cada euro invertido de media en inventario.

**Cómo interpretar el informe**  
Una tienda eficiente no es necesariamente la que menos stock tiene, sino la que consigue generar más ventas y, especialmente, más margen bruto con el inventario que mantiene.

---

### 🟢 Mejor / 🟠 Peor

En cada columna del informe, el **mejor valor válido entre las tiendas seleccionadas aparece sombreado en verde** y el **peor en naranja**. Los valores **en blanco, sin dato o iguales a cero** se excluyen de la comparación. La fila TOTAL tampoco participa en el semáforo.


**Ventas / Stock**  
🟢 **Mayor = mejor.** La tienda genera más ventas por cada euro mantenido en inventario.  
🟠 **Menor = peor.** El stock genera relativamente pocas ventas.

**Margen / Stock**  
🟢 **Mayor = mejor.** Es el indicador principal del ranking: se obtiene más margen bruto por cada euro de inventario medio.  
🟠 **Menor = peor.** El inventario está produciendo menos margen bruto.

**Rotación Coste / Stock**  
🟢 **En general, mayor = mayor rotación.** La mercancía se renueva más veces durante el periodo.  
🟠 **Muy bajo = posible exceso de stock o baja salida.**  
⚠️ Un valor excesivamente alto también debe revisarse, porque podría indicar un stock demasiado ajustado y riesgo de faltas de mercancía.

**Inventario Medio**  
No es mejor simplemente por ser más alto o más bajo. Debe analizarse junto con ventas, margen y rotación. Una tienda puede necesitar más stock porque vende mucho más.

**Ranking**  
🟢 **1 = mejor eficiencia de inventario**, según Margen / Stock.  
Cuanto mayor sea el número del ranking, menor es el margen generado por euro de inventario respecto a las demás tiendas seleccionadas.

**Importante:** los ratios deben compararse entre tiendas para el **mismo periodo seleccionado**, ya que ventas, margen y coste se acumulan durante ese periodo.
                """
            )

        render_aggrid_table(
            df_eff_disp,
            modo="auto",
            df_numericos=df_eff_num,
            resaltar_extremos_filas=True,
            columnas_extremos=[
                "Inventario Medio",
                "Ventas Periodo",
                "% Margen Acumulado",
                "Margen Bruto Periodo",
                "Coste Ventas Periodo",
                "Ventas / Stock",
                "Margen / Stock",
                "Rotación Coste / Stock",
            ],
            altura_fila=32,
            altura_cabecera=38,
            clave_preferencias="eficiencia_inventario",
        )

        descargar_excel(
            df_eff_num,
            "Eficiencia Inventario",
            f"Eficiencia_Inventario_{ano}.xlsx",
            "Descargar eficiencia en Excel",
        )
        st.stop()

    inv_periodo = df_inventario[
        (df_inventario["Año"] == ano)
        & (df_inventario["Mes"].isin(meses_sel))
    ].copy()

    if inv_periodo.empty:
        st.info(
            "No hay inventario informado para el año y los meses seleccionados."
        )
        st.stop()

    # Todas las tiendas de inventario salvo General y las que estén a cero
    # en todo el periodo solicitado.
    tiendas_inv = []

    for tienda in sorted(
        inv_periodo["Departamento"]
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    ):
        valores_tienda = pd.to_numeric(
            inv_periodo.loc[
                inv_periodo["Departamento"].astype(str).str.strip()
                == tienda,
                "Inventario",
            ],
            errors="coerce",
        ).fillna(0)

        if valores_tienda.abs().sum() > 1e-12:
            tiendas_inv.append(tienda)

    tiendas_seleccionadas_inv = st.sidebar.multiselect(
        "Selecciona tiendas de inventario",
        tiendas_inv,
        default=tiendas_inv,
    )

    if not tiendas_seleccionadas_inv or not meses_sel:
        st.info("Selecciona al menos una tienda y un mes.")
        st.stop()

    filas_display = []
    filas_excel = []

    for mes in meses_sel:
        for tienda in tiendas_seleccionadas_inv:
            fila_inv = df_inventario[
                (df_inventario["Año"] == ano)
                & (df_inventario["Mes"] == mes)
                & (
                    df_inventario["Departamento"].astype(str).str.strip()
                    == str(tienda).strip()
                )
            ]

            if fila_inv.empty:
                continue

            inventario_final = float(
                pd.to_numeric(
                    fila_inv["Inventario"], errors="coerce"
                ).fillna(0).sum()
            )

            # GENERAL representa inventario común/almacén.
            # En su fila solo se muestra el inventario final; no se calculan
            # ventas medias, margen, venta a coste ni cobertura.
            if str(tienda).strip().upper() == "GENERAL":
                venta_media_12m = None
                margen_acumulado = None
                venta_media_coste = None
                meses_stock = None
                estado = ""
            else:
                calculo = calcular_cobertura_inventario(
                    df,
                    tienda,
                    ano,
                    mes,
                    inventario_final,
                )

                venta_media_12m = calculo["venta_media_12m"]
                margen_acumulado = calculo["margen_acumulado"]
                venta_media_coste = calculo["venta_media_coste"]
                meses_stock = calculo["meses_stock"]
                estado = calculo["estado"]

            meses_stock_txt = ""
            if meses_stock is not None:
                meses_stock_txt = (
                    f"{meses_stock:,.2f}"
                    .replace(",", "X")
                    .replace(".", ",")
                    .replace("X", ".")
                )

            filas_display.append(
                {
                    "Tienda": tienda,
                    "Mes": mes,
                    "Inventario Final": formato_moneda(inventario_final),
                    "Venta Media 12M": (
                        formato_moneda(venta_media_12m)
                        if venta_media_12m is not None
                        else ""
                    ),
                    "Margen Acumulado": (
                        formato_porcentaje(margen_acumulado)
                        if margen_acumulado is not None
                        else ""
                    ),
                    "Venta Media a Coste": (
                        formato_moneda(venta_media_coste)
                        if venta_media_coste is not None
                        else ""
                    ),
                    "Meses de Stock": meses_stock_txt,
                    "Estado": estado,
                }
            )

            filas_excel.append(
                {
                    "Tienda": tienda,
                    "Año": ano,
                    "Mes": mes,
                    "Inventario Final": inventario_final,
                    "Venta Media 12M": venta_media_12m,
                    "Margen Acumulado": margen_acumulado,
                    "Venta Media a Coste": venta_media_coste,
                    "Meses de Stock": meses_stock,
                    "Estado": estado,
                }
            )

    if not filas_display:
        st.info(
            "No hay datos de inventario para las tiendas y meses seleccionados."
        )
        st.stop()

    # Añadir TOTAL por cada mes solicitado.
    # El inventario TOTAL incluye GENERAL. La venta media TOTAL se calcula
    # directamente con las ventas de toda la empresa de los últimos 12 meses.
    # El margen acumulado TOTAL se toma exclusivamente del ERP.
    df_margenes_totales = load_margenes_totales_acumulados(firma_archivo_datos())

    for mes in meses_sel:
        filas_mes = [f for f in filas_excel if f["Mes"] == mes]
        if not filas_mes:
            continue

        inventario_total = sum(
            float(f["Inventario Final"] or 0) for f in filas_mes
        )

        periodos_total = obtener_ultimos_12_periodos(ano, mes)
        ventas_totales_12m = []
        historial_total_completo = len(periodos_total) == 12

        for a, m in periodos_total:
            df_mes_total = df[
                (df["Año"] == a)
                & (df["Mes"] == m)
            ]
            ventas_rows_total = df_mes_total[
                df_mes_total["Resultados"] == "Ventas"
            ]

            if ventas_rows_total.empty:
                historial_total_completo = False
                break

            ventas_totales_12m.append(
                float(
                    pd.to_numeric(
                        ventas_rows_total["Importe D"],
                        errors="coerce",
                    ).fillna(0).sum()
                )
            )

        venta_media_total = (
            sum(ventas_totales_12m) / 12.0
            if historial_total_completo
            else None
        )

        fila_margen_total = df_margenes_totales[
            (df_margenes_totales["Año"] == ano)
            & (df_margenes_totales["Mes"] == mes)
        ]

        margen_total = None
        if not fila_margen_total.empty:
            margen_total = float(
                fila_margen_total["Margen Acumulado Total"].iloc[0]
            )

            # Normalización defensiva:
            # el ERP normalmente guarda 39,01 % como 0,3901.
            # Si alguna versión lo trae como 39,01, se convierte a 0,3901.
            if abs(margen_total) > 1.0:
                margen_total = margen_total / 100.0

        venta_coste_total = None
        meses_stock_total = None
        estado_total = ""

        if venta_media_total is None:
            estado_total = (
                "Imposible calcular media de ventas últimos 12 meses"
            )
        elif margen_total is None:
            estado_total = (
                "Imposible calcular: falta margen acumulado total del ERP"
            )
        else:
            venta_coste_total = venta_media_total * (1.0 - margen_total)
            if abs(venta_coste_total) > 1e-12:
                meses_stock_total = inventario_total / venta_coste_total
            else:
                estado_total = (
                    "Imposible calcular: venta media total a coste igual a cero"
                )

        meses_stock_total_txt = ""
        if meses_stock_total is not None:
            meses_stock_total_txt = (
                f"{meses_stock_total:,.2f}"
                .replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            )

        filas_display.append(
            {
                "Tienda": "TOTAL",
                "Mes": mes,
                "Inventario Final": formato_moneda(inventario_total),
                "Venta Media 12M": (
                    formato_moneda(venta_media_total)
                    if venta_media_total is not None else ""
                ),
                "Margen Acumulado": (
                    formato_porcentaje(margen_total)
                    if margen_total is not None else ""
                ),
                "Venta Media a Coste": (
                    formato_moneda(venta_coste_total)
                    if venta_coste_total is not None else ""
                ),
                "Meses de Stock": meses_stock_total_txt,
                "Estado": estado_total,
            }
        )

        filas_excel.append(
            {
                "Tienda": "TOTAL",
                "Año": ano,
                "Mes": mes,
                "Inventario Final": inventario_total,
                "Venta Media 12M": venta_media_total,
                "Margen Acumulado": margen_total,
                "Venta Media a Coste": venta_coste_total,
                "Meses de Stock": meses_stock_total,
                "Estado": estado_total,
            }
        )

    df_rotacion_display = pd.DataFrame(filas_display)
    df_rotacion_excel = pd.DataFrame(filas_excel)

    nombre_meses_inv = ", ".join(meses_sel)
    st.subheader(
        f"Cobertura de Inventario — {nombre_meses_inv} {ano}"
    )

    st.caption(
        "Meses de Stock = Inventario final ÷ Venta media mensual de los "
        "últimos 12 meses a coste. Para cada tienda se utiliza su margen bruto "
        "acumulado desde enero hasta el mes analizado. Para TOTAL se utiliza "
        "el margen acumulado oficial del ERP de la hoja MargenesTotalesAcumulados."
    )

    render_aggrid_table(
        df_rotacion_display,
        modo="auto",
        altura_fila=32,
        altura_cabecera=38,
        clave_preferencias="inventario_rotacion",
    )

    # Gráfico: Meses de Stock por tienda del último mes seleccionado.
    df_graf_rot = df_rotacion_excel.copy()
    if not df_graf_rot.empty:
        df_graf_rot["Meses de Stock"] = pd.to_numeric(
            df_graf_rot["Meses de Stock"], errors="coerce"
        )
        df_graf_rot = df_graf_rot[
            (df_graf_rot["Tienda"].astype(str).str.strip().str.upper() != "TOTAL")
            & df_graf_rot["Meses de Stock"].notna()
            & df_graf_rot["Meses de Stock"].ne(0)
        ].copy()

        if not df_graf_rot.empty:
            orden_meses = {m: i for i, m in enumerate(MESES_ORDEN)}
            df_graf_rot["_orden_mes"] = df_graf_rot["Mes"].map(orden_meses)
            ultimo_orden = df_graf_rot["_orden_mes"].max()
            df_graf_rot = df_graf_rot[
                df_graf_rot["_orden_mes"] == ultimo_orden
            ].copy()
            df_graf_rot = df_graf_rot.sort_values(
                "Meses de Stock", ascending=False
            )

            mes_grafico = str(df_graf_rot["Mes"].iloc[0])
            st.markdown(f"#### Meses de Stock por tienda — {mes_grafico}")
            st.caption(
                "Muestra la cobertura del último mes seleccionado. "
                "Un valor más alto indica más meses de stock disponible."
            )
            st.bar_chart(
                df_graf_rot.set_index("Tienda")["Meses de Stock"],
                use_container_width=True,
            )

    descargar_excel(
        df_rotacion_excel,
        "Rotacion",
        f"Inventario_Rotacion_{ano}.xlsx",
        "Descargar análisis de inventario en Excel",
    )

