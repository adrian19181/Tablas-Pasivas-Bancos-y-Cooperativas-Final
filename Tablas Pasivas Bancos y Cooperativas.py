import io
import os
import re
import sys
import pandas as pd
import requests
import streamlit as st

# Configuración inicial de la página
st.set_page_config(layout="wide", initial_sidebar_state="collapsed")

# Estilos CSS personalizados para tablas móviles con inmovilización 2D Sticky y ajuste inteligente de columnas
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {
        padding-top: 0rem;
        padding-bottom: 0rem;
        padding-left: 0.5rem;
        padding-right: 0.5rem;
    }

    /* Ajuste de menú desplegable (selectbox) para teclados táctiles en pantalla móvil */
    div[data-baseweb="popover"] {
        z-index: 999999 !important;
    }
    div[data-baseweb="popover"] ul,
    div[data-baseweb="popover"] div[role="listbox"],
    ul[data-baseweb="menu"] {
        max-height: 28vh !important;
        overflow-y: auto !important;
        -webkit-overflow-scrolling: touch !important;
    }
    
    /* Tarjeta y Contenedor de Tabla */
    .table-card {
        margin-bottom: 1.5rem;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    
    .table-title {
        background-color: #1e293b;
        color: #ffffff;
        font-weight: bold;
        font-size: 13px;
        text-align: center;
        padding: 8px 10px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .table-container {
        overflow-x: auto;
        overflow-y: auto;
        max-height: 420px;
        padding: 0px !important;
        margin: 0px !important;
    }
    
    /* Tabla Base */
    table.custom-table {
        border-collapse: separate;
        border-spacing: 0;
        width: max-content !important;
        min-width: 100%;
        table-layout: auto !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        font-size: 12px;
    }
    
    /* Encabezados Fijos (Sticky Top) */
    table.custom-table th {
        position: sticky !important;
        top: 0 !important;
        z-index: 20 !important;
        background-color: #334155 !important;
        color: #ffffff !important;
        font-weight: bold;
        padding: 8px 8px;
        text-align: center;
        white-space: normal !important;
        word-break: normal !important;
        width: auto !important;
        min-width: max-content !important;
        border-bottom: 1px solid #cbd5e1;
        border-right: 1px solid #cbd5e1;
    }
    
    /* Esquina Superior Izquierda Fija (Sticky Top & Left) */
    table.custom-table th:first-child {
        position: sticky !important;
        top: 0 !important;
        left: 0 !important;
        z-index: 50 !important;
        background-color: #1e293b !important;
        width: auto !important;
        min-width: max-content !important;
        white-space: normal !important;
        text-align: center;
    }
    
    /* Celdas de Datos Normales */
    table.custom-table td {
        position: static !important;
        z-index: auto !important;
        padding: 7px 10px;
        text-align: center;
        border-bottom: 1px solid #e2e8f0;
        border-right: 1px solid #e2e8f0;
        white-space: nowrap !important;
        width: auto !important;
        min-width: max-content !important;
    }
    
    /* Primera Columna Fija (Sticky Left) */
    table.custom-table td.lbl-sticky-col {
        position: sticky !important;
        left: 0 !important;
        z-index: 30 !important;
        font-weight: bold;
        text-align: left;
        width: auto !important;
        min-width: max-content !important;
        white-space: nowrap !important;
        background-clip: padding-box !important;
        box-shadow: 2px 0 5px -2px rgba(0,0,0,0.2) !important;
    }
    
    /* Filas Alternadas */
    .row-even td {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    .row-odd td {
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    .row-even td.lbl-sticky-col {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    .row-odd td.lbl-sticky-col {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("Tablas Pasivas Bancos y coopertivas")

if st.button("🔄 Refrescar"):
    st.rerun()

URL_EXCEL_DRIVE = "https://docs.google.com/spreadsheets/d/1kiwAirOeaVEYxOdRP3FWB2xH-IewYOcx/export?format=xlsx"


@st.cache_data(ttl=300)
def cargar_datos_excel(url):
    response = requests.get(url)
    excel_file = io.BytesIO(response.content)
    return pd.read_excel(excel_file, sheet_name=None, header=None)


def limpiar_texto_encabezado(texto):
    t = str(texto).strip()
    # Eliminación de texto redundante como "PERSONAS NATURALES Y JURIDICAS NO FINANCIERAS"
    t = re.sub(
        r"PERSONAS\s+NATURALES\s+Y\s+JURÍDICAS\s+NO\s+FINANCIERAS",
        "",
        t,
        flags=re.IGNORECASE,
    ).strip()
    t = re.sub(
        r"PERSONAS\s+NATURALES\s+Y\s+JURIDICAS\s+NO\s+FINANCIERAS",
        "",
        t,
        flags=re.IGNORECASE,
    ).strip()

    # Estandarización de nomenclaturas a TNA y TEA
    t = re.sub(
        r"Tasa\s+de\s+interés\s+nominal\s+anual", "TNA", t, flags=re.IGNORECASE
    )
    t = re.sub(
        r"Tasa\s+de\s+interés\s+efectiva\s+anual", "TEA", t, flags=re.IGNORECASE
    )
    t = re.sub(r"Tasa\s+nominal\s+anual", "TNA", t, flags=re.IGNORECASE)
    t = re.sub(r"Tasa\s+efectiva\s+anual", "TEA", t, flags=re.IGNORECASE)
    t = re.sub(r"Tasa\s+Nominal\b", "TNA", t, flags=re.IGNORECASE)
    t = re.sub(r"Tasa\s+Efectiva\b", "TEA", t, flags=re.IGNORECASE)
    t = re.sub(r"\bNominal\b", "TNA", t, flags=re.IGNORECASE)
    t = re.sub(r"\bEfectiva\b", "TEA", t, flags=re.IGNORECASE)

    t = re.sub(r"\s+", " ", t).strip()
    return t


def detectar_unidad_primera_columna(header_original, rows):
    h_str = str(header_original).strip()

    if re.search(
        r"\((?:\$|días|dias|meses|mes|Días|Dias|Meses)\)", h_str, re.IGNORECASE
    ):
        return h_str

    h_upper = h_str.upper()
    if "FECHA" in h_upper:
        return h_str

    vals = [
        str(r[0]).strip().upper() for r in rows if len(r) > 0 and pd.notna(r[0])
    ]
    texto_unido = " ".join(vals)

    if (
        "$" in texto_unido
        or "DÓLAR" in texto_unido
        or "DOLAR" in texto_unido
        or "$" in h_upper
        or "DÓLAR" in h_upper
        or "DOLAR" in h_upper
        or "MONTO" in h_upper
        or "SALDO" in h_upper
        or "RANGOS" == h_upper
    ):
        unit = "($)"
    elif "MES" in texto_unido or "MESES" in h_upper:
        unit = "(Meses)"
    elif (
        "DÍA" in texto_unido
        or "DIA" in texto_unido
        or "DÍA" in h_upper
        or "DIA" in h_upper
        or "PLAZO" in h_upper
        or "DESDE" in h_upper
    ):
        unit = "(Días)"
    elif "RANGO" in h_upper:
        unit = "($)"
    else:
        unit = ""

    if any(k in h_upper for k in ["DÍAS", "DIAS", "DÓLARES", "DOLARES", "MESES"]):
        if ("DÓLARES" in h_upper or "DOLARES" in h_upper) and "($)" not in h_str:
            return f"{h_str} ($)"
        return h_str

    if unit and unit.lower() not in h_str.lower():
        return f"{h_str} {unit}"
    return h_str


def extraer_subtablas(df_raw):
    num_cols = df_raw.shape[1]
    col_blocks = []
    current_block = []

    for c in range(num_cols):
        val0 = df_raw.iloc[0, c]
        val1 = df_raw.iloc[1, c]
        is_val = pd.notna(val0) or pd.notna(val1)
        if is_val:
            current_block.append(c)
        else:
            if current_block:
                col_blocks.append(current_block)
                current_block = []
    if current_block:
        col_blocks.append(current_block)

    subtablas = []
    for block in col_blocks:
        titulos = [
            str(df_raw.iloc[0, c]).strip()
            for c in block
            if pd.notna(df_raw.iloc[0, c])
        ]
        titulo = titulos[0] if titulos else "TABLA DE DATOS"
        headers = [
            (
                limpiar_texto_encabezado(df_raw.iloc[1, c])
                if pd.notna(df_raw.iloc[1, c])
                else f"Col {i+1}"
            )
            for i, c in enumerate(block)
        ]
        df_sub = df_raw.iloc[2:, block].dropna(how="all")
        rows = df_sub.values.tolist()

        subtabla = {"titulo": titulo, "headers": headers, "rows": rows}

        # Eliminación de columna 'Cta. de Ahorros JEPito' en Cooperativa JEP
        indices_validos = [
            i
            for i, h in enumerate(subtabla["headers"])
            if "jepito" not in str(h).lower()
        ]
        if len(indices_validos) < len(subtabla["headers"]):
            subtabla["headers"] = [subtabla["headers"][i] for i in indices_validos]
            subtabla["rows"] = [
                [r[i] for i in indices_validos if i < len(r)]
                for r in subtabla["rows"]
            ]

        # Unificación Produbanco: 'Plazo desde' y 'Plazo hasta'
        if (
            len(subtabla["headers"]) >= 2
            and "PLAZO DESDE" in str(subtabla["headers"][0]).upper()
            and "PLAZO HASTA" in str(subtabla["headers"][1]).upper()
        ):
            nuevos_headers = ["Rango de Días"] + subtabla["headers"][2:]
            nuevas_rows = []
            for r in subtabla["rows"]:
                p_desde = str(r[0]).replace(".0", "") if pd.notna(r[0]) else ""
                p_hasta = str(r[1]).replace(".0", "") if pd.notna(r[1]) else ""
                rango_str = (
                    f"De {p_desde} a {p_hasta} días"
                    if p_desde and p_hasta
                    else p_desde or p_hasta
                )
                nuevas_rows.append([rango_str] + r[2:])

            subtabla = {
                "titulo": subtabla["titulo"],
                "headers": nuevos_headers,
                "rows": nuevas_rows,
            }

        # Depuración de columna redundante inicial (Ej. Jardín Azuayo 'Tipos de Ahorros')
        headers_clean = subtabla["headers"]
        rows_clean = subtabla["rows"]

        if len(headers_clean) > 1 and str(headers_clean[0]).strip().lower() in [
            "tipos de ahorros",
            "tipo de ahorro",
            "tipos de ahorro",
        ]:
            headers_clean = headers_clean[1:]
            rows_clean = [r[1:] for r in rows_clean]
        elif len(headers_clean) > 1 and len(rows_clean) > 0:
            val0_set = set(
                str(r[0]).strip().lower()
                for r in rows_clean
                if len(r) > 0 and pd.notna(r[0])
            )
            if (
                len(val0_set) == 1
                and list(val0_set)[0] == subtabla["titulo"].strip().lower()
            ):
                headers_clean = headers_clean[1:]
                rows_clean = [r[1:] for r in rows_clean]

        subtabla["headers"] = headers_clean
        subtabla["rows"] = rows_clean

        # Detectar y asignar automáticamente unidad ($), (Días), (Meses)
        subtabla["headers"][0] = detectar_unidad_primera_columna(
            subtabla["headers"][0], subtabla["rows"]
        )

        subtablas.append(subtabla)
    return subtablas


def clasificar_producto(subtabla):
    titulo_upper = str(subtabla.get("titulo", "")).upper()
    headers_upper = " ".join(
        [str(h) for h in subtabla.get("headers", [])]
    ).upper()
    texto_completo = titulo_upper + " " + headers_upper

    tipos = []

    kw_plazo = [
        "PLAZO",
        "PÓLIZA",
        "POLIZA",
        "CERTIFICADO DE DEPÓSITO",
        "CERTIFICADOS DE DEPÓSITO",
        "CERTIFICADO",
        "ACUMULACIÓN",
        "ACUMULACION",
        "INVERSION",
        "INVERSIONES",
        "DEP. A PLAZO",
    ]
    kw_prog = [
        "PROGRAMADO",
        "FLEXIAHORRO",
        "FONDO AHORRO",
        "FONDOJEP",
        "FONDO JEP",
        "METAHORRO",
        "FUTURO SEGURO",
        "CHANCHITO",
        "CESANTÍA",
        "CESANTIA",
    ]
    kw_ahorro = [
        "AHORRO",
        "VISTA",
        "DIGITAL",
        "CUENTA",
        "SOCIOS",
        "ANGELITOS",
        "RENTABILIDAD",
        "CTA",
        "OTRAS CUENTAS",
        "REFERENCIAL",
    ]

    # Evaluamos Ahorro Programado
    es_prog = any(k in texto_completo for k in kw_prog)
    if es_prog:
        tipos.append("Ahorros Programados")

    # Evaluamos Plazos Fijos
    if any(k in titulo_upper for k in kw_plazo):
        tipos.append("Plazos Fijos")

    # Evaluamos Cuentas de Ahorros
    if any(k in texto_completo for k in kw_ahorro):
        es_plazo_puro = any(
            k in titulo_upper
            for k in [
                "DEPÓSITO A PLAZO",
                "DEPÓSITOS A PLAZO",
                "PLAZO FIJO",
                "PÓLIZA",
                "POLIZA",
                "CERTIFICADO",
            ]
        ) and not any(k in headers_upper for k in ["CTA", "CUENTA", "FONDOJEP"])

        # Si es una tabla dedicada exclusivamente a ahorro programado (y no la multifunción de JEP), no incluir en Cuentas de Ahorros
        es_prog_puro = (
            es_prog
            and "TASAS DE INTERÉS POR PRODUCTOS DE AHORRO" not in titulo_upper
        )

        if not es_plazo_puro and not es_prog_puro:
            tipos.append("Cuentas de Ahorros")

    if not tipos:
        tipos.append("Otros")

    return list(set(tipos))


def filtrar_columnas_subtabla(subtabla, producto_seleccionado):
    titulo_upper = str(subtabla.get("titulo", "")).upper()
    headers = subtabla.get("headers", [])

    if "TASAS DE INTERÉS POR PRODUCTOS DE AHORRO" in titulo_upper:
        if producto_seleccionado == "Ahorros Programados":
            cols_deseadas = [0]
            for idx, h in enumerate(headers):
                if idx > 0 and any(
                    k in str(h).upper() for k in ["FONDOJEP", "INTELIJEP"]
                ):
                    cols_deseadas.append(idx)

            nuevos_headers = [headers[i] for i in cols_deseadas]
            nuevas_rows = [[r[i] for i in cols_deseadas] for r in subtabla["rows"]]

            return {
                "titulo": subtabla["titulo"],
                "headers": nuevos_headers,
                "rows": nuevas_rows,
            }
        elif producto_seleccionado == "Cuentas de Ahorros":
            cols_deseadas = [0]
            for idx, h in enumerate(headers):
                if idx > 0 and not any(
                    k in str(h).upper() for k in ["FONDOJEP", "INTELIJEP"]
                ):
                    cols_deseadas.append(idx)

            nuevos_headers = [headers[i] for i in cols_deseadas]

            # Filtrar únicamente las filas de 'Monto' e 'Interés'
            nuevas_rows = []
            for r in subtabla["rows"]:
                detalle_val = str(r[0]).strip().upper() if pd.notna(r[0]) else ""
                if detalle_val in ["MONTO", "INTERÉS", "INTERES"]:
                    nuevas_rows.append([r[i] for i in cols_deseadas])

            return {
                "titulo": subtabla["titulo"],
                "headers": nuevos_headers,
                "rows": nuevas_rows,
            }

    return subtabla


def dividir_encabezado_dos_lineas(texto):
    t = str(texto).strip()
    if "<br>" in t or "<br/>" in t:
        return t
    palabras = t.split()
    if len(palabras) <= 1:
        return t
    mitad = len(palabras) // 2
    linea1 = " ".join(palabras[:mitad])
    linea2 = " ".join(palabras[mitad:])
    return f"{linea1}<br>{linea2}"


def formatear_encabezado(header_text):
    # Formato a dos líneas
    header_text = dividir_encabezado_dos_lineas(header_text)
    header_text = re.sub(
        r"([\$\d\.,\w]+)\s*-\s*([\$\d\.,\w]+)",
        r'<span style="white-space: nowrap;">\1&#8209;\2</span>',
        header_text,
    )
    header_text = re.sub(
        r"(De\s+[\$\d\.,]+)\s+(a\s+[\$\d\.,]+)",
        r'<span style="white-space: nowrap;">\1&nbsp;\2</span>',
        header_text,
        flags=re.IGNORECASE,
    )
    return header_text


def renderizar_tabla_html(subtabla):
    html = f'<div class="table-card">'
    html += f'<div class="table-title">{subtabla["titulo"]}</div>'
    html += '<div class="table-container">'
    html += '<table class="custom-table">'

    html += "<thead><tr>"
    for idx, header in enumerate(subtabla["headers"]):
        header_fmt = formatear_encabezado(header)
        html += f"<th>{header_fmt}</th>"
    html += "</tr></thead>"

    html += "<tbody>"
    for r_idx, row in enumerate(subtabla["rows"]):
        row_class = "row-even" if r_idx % 2 == 0 else "row-odd"
        html += f'<tr class="{row_class}">'
        for c_idx, val in enumerate(row):
            cell_val = str(val) if pd.notna(val) else ""
            if c_idx == 0:
                html += f'<td class="lbl-sticky-col">{cell_val}</td>'
            else:
                html += f"<td>{cell_val}</td>"
        html += "</tr>"
    html += "</tbody></table></div></div>"
    return html


# Carga de datos
datos_excel = cargar_datos_excel(URL_EXCEL_DRIVE)

# Mapeo para mostrar "Tasas Referenciales" en las opciones del selector
mapa_instituciones = {}
for sheet_name in datos_excel.keys():
    display_name = (
        "Tasas Referenciales"
        if sheet_name == "Tasa Pasiva Referencial"
        else sheet_name
    )
    mapa_instituciones[display_name] = sheet_name

# Ordenamiento alfabético explícito de la lista para visibilidad directa en la lista deplegable
lista_instituciones = sorted(list(mapa_instituciones.keys()))

# Filtros de entrada
institucion_seleccionada = st.selectbox(
    "Selecciona o escribe la Institución:",
    options=lista_instituciones,
    index=0 if lista_instituciones else None,
)

producto_seleccionado = st.selectbox(
    "Selecciona el Tipo de Producto:",
    options=["Todos", "Plazos Fijos", "Ahorros Programados", "Cuentas de Ahorros"],
)

# Renderizado de tablas
if institucion_seleccionada and institucion_seleccionada in mapa_instituciones:
    sheet_real = mapa_instituciones[institucion_seleccionada]
    df_hoja = datos_excel[sheet_real]
    subtablas = extraer_subtablas(df_hoja)

    tablas_filtradas = []
    for st_item in subtablas:
        tipos = clasificar_producto(st_item)
        if (
            producto_seleccionado == "Todos"
            or producto_seleccionado in tipos
            or institucion_seleccionada == "Tasas Referenciales"
        ):
            subtabla_filtrada = filtrar_columnas_subtabla(
                st_item, producto_seleccionado
            )
            tablas_filtradas.append(subtabla_filtrada)

    if tablas_filtradas:
        for st_item in tablas_filtradas:
            st.markdown(renderizar_tabla_html(st_item), unsafe_allow_html=True)
    else:
        st.info("No se encontraron tablas para el tipo de producto seleccionado.")

if __name__ == "__main__":
    if "streamlit" not in sys.argv[0]:
        os.system(f'streamlit run "{__file__}" --server.port 8501')