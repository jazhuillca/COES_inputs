# -*- coding: utf-8 -*-
"""
App Streamlit para descargar el reporte de Mantenimientos del COES
(https://www.coes.org.pe/Portal/eventos/mantenimiento/) y filtrarlo por:
Tipo de Equipo, Fecha desde/hasta, Mantenimiento (estado), Tipo de Mantenimiento,
Con Interrupción, Tipo de Empresa y Empresa.

CÓMO FUNCIONA EL FILTRADO
-------------------------
El formulario web de COES filtra en el servidor usando códigos numéricos internos
(por ejemplo tiposEquipo=41,42,44,...) que no están documentados públicamente y que
no se pueden confirmar sin inspeccionar el JavaScript interno del portal. Adivinar
esos códigos es riesgoso: si uno está mal, el reporte puede salir vacío o incompleto
sin que te des cuenta.

Por eso esta app usa un enfoque más seguro:
  1) Se pide SIEMPRE el reporte completo al servidor (con todos los tipos
     seleccionados, igual que el formulario por defecto), solo variando el rango
     de fechas y el estado de "Mantenimiento" (Ejecutados / Programado Diario /
     Programado Semanal / Programado Mensual), que sí están confirmados.
  2) Una vez descargado el Excel, el filtrado por Tipo de Equipo, Tipo de
     Mantenimiento, Con Interrupción, Tipo de Empresa y Empresa se hace en
     pandas, sobre las columnas reales del archivo (texto, no códigos).

Esto es más lento (siempre se trae todo el rango de fechas) pero es exacto.

Ejecutar con:  streamlit run coes_mantenimientos_app.py
"""

import io
import requests
import pandas as pd
import streamlit as st
from datetime import date, timedelta

# =========================================================================
# CONFIG
# =========================================================================
st.set_page_config(page_title="COES - Consulta de Mantenimientos", layout="wide")
st.title("🔧 Consulta de Mantenimientos - COES")

BASE = "https://www.coes.org.pe/Portal/eventos"
REFERER = "https://www.coes.org.pe/Portal/eventos/mantenimiento/"

# --- Valores "todos seleccionados" que se envían siempre al servidor ---
DEFAULT_TIPOS_EQUIPO = (
    "41,42,44,53,45,46,51,49,50,52,55,43,47,48,54,-1,0,1,2,3,4,5,6,7,8,9,10,"
    "11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,"
    "35,40,36,37,38,39,56"
)
DEFAULT_TIPOS_MANTTO = "1,2,3,4,5,6,7,9,10,12"
DEFAULT_TIPOS_EMPRESA_CODES = "1,2,3,4,5"

# --- Estado del mantenimiento: confirmado en el script original que 1=EJECUTADOS.
# Los otros 3 valores (2,3,4) se asumen por el orden del <select> del formulario;
# si alguno no coincide, el propio reporte descargado lo evidenciará (columnas
# vacías o distintas al filtro elegido) y podrás avisarme para corregirlo.
MANTENIMIENTO_ESTADOS = {
    "EJECUTADOS": "1",
    "PROGRAMADO DIARIO": "2",
    "PROGRAMADO SEMANAL": "3",
    "PROGRAMADO MENSUAL": "4",
}

# --- Listas de texto reales, tal como aparecen en el formulario web (para los
# filtros que aplicamos en pandas sobre el archivo ya descargado) ---
TIPOS_EQUIPO_LABELS = [
    "CUENCA HIDROLOGICA", "ESTACIÓN HIDROLÓGICA", "PULVERIZADOR", "RELE PROTECCIONES",
    "SUMINISTRO", "STATCOM", "TRAMPA DE ONDA", "BESS", "PARTICIPANTE - MME",
    "CONJUNTO DE LINEAS", "ESQ. DE CONTROL Y PROTECCIÓN", "BLACK START",
    "TRANSFORMADOR 4D", "TRANSFORMADOR 5D", "COMPONENTE BESS", "(NO DEFINIDO)",
    "SUBESTACION", "GENERADOR HIDROELÉCTRICO", "GENERADOR TERMOELÉCTRICO",
    "CENTRAL HIDROELÉCTRICA", "CENTRAL TERMOELÉCTRICA", "CELDA", "BARRA",
    "LINEA DE TRANSMISION", "TRANSFORMADOR 2D", "TRANSFORMADOR 3D",
    "BANCO DE CONDENSADOR PARALELO", "REACTOR", "COMPENSADOR SINCRONO", "SVC",
    "BANCO DE CONDENSADORES SERIE", "INTERRUPTOR", "ACOPLAMIENTO", "TUBERIA",
    "EMBALSE:PRESA,TAZA,PULMON", "CANAL,TOMA,DESCARGA", "SISTEMA DE BARRAS (SSEE)",
    "CALDERO", "RIO", "SECCIONADOR", "RELE", "TRANSFORMADOR DE CORRIENTE",
    "TRANSFORMADOR DE TENSION", "PARARRAYOS", "TRANSFORMADOR ZIG-ZAG",
    "MOTOR SINCRONO", "MOTOR ASINCRONO", "CLIENTE", "PLANTA", "GASEODUCTO",
    "CARGA", "SERVICIOS AUXILIARES", "GENERADOR SOLAR", "CENTRAL SOLAR",
    "GENERADOR EOLICO", "CENTRAL EOLICA",
]

# Tipos de equipo de generación (los 4 que pediste antes) preseleccionados por defecto
GENERADORES_DEFAULT = [
    "GENERADOR HIDROELÉCTRICO",
    "GENERADOR SOLAR",
    "GENERADOR TERMOELÉCTRICO",
    "GENERADOR EOLICO",
]

TIPOS_MANTTO_LABELS = [
    "MANTENIMIENTO PREVENTIVO", "MANTENIMIENTO CORRECTIVO", "AMPLIACION Y/O MEJORAS",
    "EVENTO", "FALLA", "PRUEBAS", "FALLA ENVIO ICCP", "OTROS",
    "ENERGIZACION DE NUEVOS EQUIPOS O INSTALACIONES", "SEGURIDAD DE LAS PERSONAS",
]

TIPOS_EMPRESA_LABELS = [
    "TRANSMISION", "DISTRIBUCION", "GENERACION", "USUARIO LIBRE", "FALTA DEFINIR",
]

INTERRUPCION_OPCIONES = ["--TODOS--", "SI", "NO"]


# =========================================================================
# DESCARGA DEL REPORTE (server-side: solo fecha + estado de mantenimiento)
# =========================================================================
def _session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": REFERER,
        "Origin": "https://www.coes.org.pe",
    })
    s.get(REFERER, timeout=30)
    return s


@st.cache_data(show_spinner=False, ttl=3600)
def descargar_reporte_mantenimientos(fecha_inicial: str, fecha_final: str, tipo_mantenimiento_code: str) -> bytes:
    """
    fecha_inicial / fecha_final: strings dd/mm/yyyy.
    tipo_mantenimiento_code: código del estado (ver MANTENIMIENTO_ESTADOS).
    Devuelve los bytes crudos del Excel descargado.
    """
    s = _session()

    payload = {
        "tiposMantenimiento": tipo_mantenimiento_code,
        "fechaInicial": fecha_inicial,
        "fechaFinal": fecha_final,
        "indispo": "-1",
        "tiposEmpresa": DEFAULT_TIPOS_EMPRESA_CODES,
        "empresas": "-1",
        "tiposEquipo": DEFAULT_TIPOS_EQUIPO,
        "interrupcion": "-1",
        "tiposMantto": DEFAULT_TIPOS_MANTTO,
    }

    resp = s.post(f"{BASE}/mantenimiento/GenerarArchivoReporte", data=payload, timeout=90)
    resp.raise_for_status()

    try:
        resultado = resp.json()
    except ValueError:
        resultado = resp.text.strip()

    if str(resultado) != "1":
        raise RuntimeError(f"El servidor de COES no confirmó éxito al generar el reporte (respuesta: {resultado!r}).")

    descarga = s.get(f"{BASE}/mantenimiento/ExportarReporte", params={"tipo": 0}, timeout=90)
    descarga.raise_for_status()

    content_type = descarga.headers.get("Content-Type", "")
    if "spreadsheet" not in content_type and "excel" not in content_type and "octet-stream" not in content_type:
        raise RuntimeError(
            f"COES no devolvió un Excel válido (Content-Type={content_type!r}). "
            f"Puede que el rango de fechas sea muy amplio o que la sesión haya expirado."
        )

    return descarga.content


# =========================================================================
# DETECCIÓN AUTOMÁTICA DE COLUMNAS EN EL EXCEL DESCARGADO
# =========================================================================
def detectar_columna(df: pd.DataFrame, keywords):
    """Busca la primera columna cuyo nombre contenga TODAS las keywords dadas (case-insensitive)."""
    for col in df.columns:
        nombre = str(col).upper()
        if all(kw.upper() in nombre for kw in keywords):
            return col
    return None


# =========================================================================
# UI - FILTROS
# =========================================================================
with st.sidebar:
    st.header("📅 Rango de fechas (obligatorio)")
    hoy = date.today()
    fecha_desde = st.date_input("Fecha desde", value=hoy - timedelta(days=30), format="DD/MM/YYYY")
    fecha_hasta = st.date_input("Fecha hasta", value=hoy, format="DD/MM/YYYY")

    st.header("⚙️ Estado del Mantenimiento")
    mantenimiento_sel = st.selectbox(
        "Mantenimiento",
        options=list(MANTENIMIENTO_ESTADOS.keys()),
        index=0,
        help="Confirmado: EJECUTADOS. Los otros 3 estados se asumen por orden del formulario web.",
    )

    st.header("🔩 Filtros (aplicados sobre el archivo descargado)")
    tipo_equipo_sel = st.multiselect(
        "Tipo de Equipo", options=TIPOS_EQUIPO_LABELS, default=GENERADORES_DEFAULT
    )
    tipo_mantto_sel = st.multiselect(
        "Tipo de Mantenimiento", options=TIPOS_MANTTO_LABELS, default=[]
    )
    interrupcion_sel = st.selectbox("Con Interrupción", options=INTERRUPCION_OPCIONES, index=0)
    tipo_empresa_sel = st.multiselect(
        "Tipo de Empresa", options=TIPOS_EMPRESA_LABELS, default=[]
    )
    empresa_busqueda = st.text_input(
        "Empresa (buscar texto parcial, opcional)",
        help="La lista de empresas en COES tiene miles de entradas; busca por texto en vez de seleccionar de una lista.",
    )

    st.divider()
    buscar = st.button("🚀 Descargar y filtrar", type="primary", use_container_width=True)


# =========================================================================
# EJECUCIÓN
# =========================================================================
# --- Persistimos los parámetros de la última búsqueda en session_state.
# Esto es clave: al presionar un botón de descarga, Streamlit vuelve a ejecutar
# todo el script y el botón "buscar" deja de estar presionado (vuelve a False).
# Si el bloque de resultados dependiera solo de `buscar`, desaparecería justo
# antes de que el navegador complete la descarga. Con session_state, el bloque
# se sigue mostrando (y descargando_reporte_mantenimientos usa caché, así que
# no se vuelve a golpear el servidor de COES en cada rerun).
if buscar:
    if fecha_desde > fecha_hasta:
        st.error("La fecha 'desde' no puede ser posterior a la fecha 'hasta'.")
        st.stop()
    st.session_state["coes_fetch_params"] = (
        fecha_desde.strftime("%d/%m/%Y"),
        fecha_hasta.strftime("%d/%m/%Y"),
        MANTENIMIENTO_ESTADOS[mantenimiento_sel],
        mantenimiento_sel,
    )

if "coes_fetch_params" in st.session_state:
    fi, ff, code, mantenimiento_label = st.session_state["coes_fetch_params"]

    with st.spinner(f"Descargando reporte de COES ({fi} a {ff}, {mantenimiento_label})..."):
        try:
            raw_bytes = descargar_reporte_mantenimientos(fi, ff, code)
        except Exception as e:
            st.error(f"❌ Error al descargar el reporte: {e}")
            st.stop()

    try:
        df = pd.read_excel(io.BytesIO(raw_bytes))
    except Exception as e:
        st.error(
            f"❌ El archivo descargado no se pudo leer como Excel ({e}). "
            f"Puede que COES haya devuelto una página de error en vez del reporte."
        )
        st.stop()

    df = df.dropna(how="all").reset_index(drop=True)
    st.success(f"✅ Reporte descargado: {len(df):,} filas, {len(df.columns)} columnas.")

    with st.expander("👀 Ver columnas originales del reporte (para verificar auto-detección)"):
        st.write(list(df.columns))

    # --- Auto-detección de columnas relevantes ---
    col_equipo = detectar_columna(df, ["TIPO", "EQUIPO"])
    col_mantto = detectar_columna(df, ["TIPO", "MANTENIMIENTO"]) or detectar_columna(df, ["TIPO", "MANTTO"])
    col_interrup = detectar_columna(df, ["INTERRUP"])
    col_tipo_empresa = detectar_columna(df, ["TIPO", "EMPRESA"])
    col_empresa = detectar_columna(df, ["EMPRESA"])
    # si "TIPO EMPRESA" y "EMPRESA" apuntaron a la misma columna, buscar otra para empresa
    if col_empresa == col_tipo_empresa:
        candidatos = [c for c in df.columns if "EMPRESA" in str(c).upper() and c != col_tipo_empresa]
        col_empresa = candidatos[0] if candidatos else col_empresa

    st.subheader("🔎 Verificación de columnas detectadas")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        col_equipo = st.selectbox("Col. Tipo de Equipo", options=[None] + list(df.columns),
                                   index=(list(df.columns).index(col_equipo) + 1) if col_equipo in df.columns else 0)
    with c2:
        col_mantto = st.selectbox("Col. Tipo de Mantenimiento", options=[None] + list(df.columns),
                                   index=(list(df.columns).index(col_mantto) + 1) if col_mantto in df.columns else 0)
    with c3:
        col_interrup = st.selectbox("Col. Con Interrupción", options=[None] + list(df.columns),
                                     index=(list(df.columns).index(col_interrup) + 1) if col_interrup in df.columns else 0)
    with c4:
        col_tipo_empresa = st.selectbox("Col. Tipo de Empresa", options=[None] + list(df.columns),
                                         index=(list(df.columns).index(col_tipo_empresa) + 1) if col_tipo_empresa in df.columns else 0)
    with c5:
        col_empresa = st.selectbox("Col. Empresa", options=[None] + list(df.columns),
                                    index=(list(df.columns).index(col_empresa) + 1) if col_empresa in df.columns else 0)

    # --- Aplicar filtros ---
    df_filtrado = df.copy()

    if tipo_equipo_sel and col_equipo:
        df_filtrado = df_filtrado[
            df_filtrado[col_equipo].astype(str).str.upper().str.strip().isin([t.upper() for t in tipo_equipo_sel])
        ]

    if tipo_mantto_sel and col_mantto:
        df_filtrado = df_filtrado[
            df_filtrado[col_mantto].astype(str).str.upper().str.strip().isin([t.upper() for t in tipo_mantto_sel])
        ]

    if interrupcion_sel != "--TODOS--" and col_interrup:
        df_filtrado = df_filtrado[
            df_filtrado[col_interrup].astype(str).str.upper().str.strip() == interrupcion_sel.upper()
        ]

    if tipo_empresa_sel and col_tipo_empresa:
        df_filtrado = df_filtrado[
            df_filtrado[col_tipo_empresa].astype(str).str.upper().str.strip().isin([t.upper() for t in tipo_empresa_sel])
        ]

    if empresa_busqueda.strip() and col_empresa:
        df_filtrado = df_filtrado[
            df_filtrado[col_empresa].astype(str).str.upper().str.contains(empresa_busqueda.strip().upper(), na=False)
        ]

    df_filtrado = df_filtrado.reset_index(drop=True)

    st.subheader(f"📋 Resultado filtrado: {len(df_filtrado):,} de {len(df):,} filas")
    st.dataframe(df_filtrado, use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        buf_xlsx = io.BytesIO()
        df_filtrado.to_excel(buf_xlsx, index=False, engine="openpyxl")
        st.download_button(
            "⬇️ Descargar filtrado (.xlsx)",
            data=buf_xlsx.getvalue(),
            file_name="coes_mantenimientos_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_b:
        buf_csv = df_filtrado.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Descargar filtrado (.csv)",
            data=buf_csv,
            file_name="coes_mantenimientos_filtrado.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with st.expander("📦 Descargar también el reporte SIN filtrar (original de COES)"):
        st.download_button(
            "⬇️ Descargar original (.xlsx)",
            data=raw_bytes,
            file_name="coes_mantenimientos_original.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Configura los filtros en el panel izquierdo y presiona **Descargar y filtrar**.")

if "coes_fetch_params" in st.session_state:
    if st.sidebar.button("🗑️ Limpiar resultado / nueva búsqueda"):
        del st.session_state["coes_fetch_params"]
        st.rerun()
