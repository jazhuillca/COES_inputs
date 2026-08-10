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
import time
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

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


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
def descargar_reporte_mantenimientos(
    fecha_inicial: str, fecha_final: str, tipo_mantenimiento_code: str, intento: int = 0
) -> bytes:
    """
    fecha_inicial / fecha_final: strings dd/mm/yyyy.
    tipo_mantenimiento_code: código del estado (ver MANTENIMIENTO_ESTADOS).
    intento: no se usa en la lógica, solo sirve para invalidar el caché de
    Streamlit cuando se reintenta la misma descarga (ver descargar_y_parsear_reporte).
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

    # El servidor genera el archivo de forma asíncrona: si se pide la
    # descarga inmediatamente después, a veces todavía no terminó de generar
    # el reporte nuevo y devuelve el último que sí tenía listo (de una
    # consulta anterior, con otro rango de fechas). Esta pequeña espera le da
    # tiempo a terminar antes de exportar.
    time.sleep(1.5)

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


def _generar_rangos_anuales(fecha_desde: date, fecha_hasta: date):
    """Divide un rango de fechas en sub-rangos de máximo 1 año calendario cada uno.

    Se comprobó que el portal de COES trunca/limita silenciosamente los reportes
    que abarcan un rango muy amplio (p.ej. pedir 01/01/2024 a 31/07/2026 solo
    devuelve datos de 2026, sin avisar del recorte). Pidiendo año por año y
    combinando los resultados en pandas se evita ese límite del servidor.
    Devuelve una lista de tuplas (desde, hasta) como objetos date.
    """
    rangos = []
    inicio_actual = fecha_desde
    while inicio_actual <= fecha_hasta:
        fin_de_anio = date(inicio_actual.year, 12, 31)
        fin_actual = min(fin_de_anio, fecha_hasta)
        rangos.append((inicio_actual, fin_actual))
        inicio_actual = fin_actual + timedelta(days=1)
    return rangos


def _parsear_excel_coes(raw_bytes: bytes) -> pd.DataFrame:
    """Convierte los bytes crudos de un Excel de COES en un DataFrame limpio:
    detecta la fila real de encabezados y descarta columnas líder vacías."""
    fila_header = detectar_fila_encabezado(raw_bytes)
    df_parte = pd.read_excel(io.BytesIO(raw_bytes), header=fila_header)
    while (
        len(df_parte.columns) > 0
        and str(df_parte.columns[0]).startswith("Unnamed")
        and df_parte[df_parte.columns[0]].isna().all()
    ):
        df_parte = df_parte.drop(columns=df_parte.columns[0])
    return df_parte.dropna(how="all").reset_index(drop=True)


def _validar_rango_chunk(df_parte: pd.DataFrame, desde_i: date, hasta_i: date) -> bool:
    """Verifica que un chunk descargado realmente tenga fechas dentro del rango
    pedido. Esto detecta el caso en que COES devuelve un reporte "viejo" (de una
    consulta anterior) porque el archivo nuevo todavía no terminó de generarse
    del lado del servidor. Si no se encuentra una columna de fecha reconocible,
    se asume OK (no se puede verificar, mejor no bloquear falsos negativos)."""
    col_fecha = detectar_columna(df_parte, ["INICIO"])
    if col_fecha is None or df_parte.empty:
        return True
    fechas = pd.to_datetime(df_parte[col_fecha], errors="coerce", dayfirst=True).dropna()
    if fechas.empty:
        return True
    return bool(((fechas.dt.date >= desde_i) & (fechas.dt.date <= hasta_i)).any())


def descargar_y_parsear_reporte(
    fecha_desde: date, fecha_hasta: date, tipo_mantenimiento_code: str, progreso=None, estado_texto=None
) -> tuple[pd.DataFrame, list]:
    """Descarga el reporte completo pidiéndolo año por año (ver _generar_rangos_anuales)
    y devuelve (df_combinado, errores). `errores` es una lista de strings describiendo
    qué tramos anuales fallaron (si alguno falla, se sigue con los demás en vez de
    abortar toda la descarga). `progreso`, si se pasa, es un st.progress() que se va
    actualizando por cada sub-rango descargado; `estado_texto`, si se pasa, es un
    st.empty() donde se va mostrando qué año se está pidiendo.

    Cada tramo se valida contra el rango pedido (_validar_rango_chunk): si COES
    devuelve un reporte que no corresponde (p.ej. quedó un reporte anterior sin
    terminar de regenerarse del lado del servidor), se reintenta un par de veces
    forzando una descarga nueva (evitando el caché local de Streamlit) antes de
    darlo por fallido."""
    rangos = _generar_rangos_anuales(fecha_desde, fecha_hasta)
    MAX_INTENTOS = 3
    partes = []
    errores = []
    for i, (desde_i, hasta_i) in enumerate(rangos, start=1):
        etiqueta_anio = desde_i.year if desde_i.year == hasta_i.year else f"{desde_i.year}-{hasta_i.year}"
        df_parte = None
        ultimo_error = None
        for intento in range(MAX_INTENTOS):
            if estado_texto is not None:
                sufijo = f" (reintento {intento}/{MAX_INTENTOS - 1})" if intento else ""
                estado_texto.text(f"Descargando tramo {i}/{len(rangos)} ({etiqueta_anio}){sufijo}...")
            try:
                raw_bytes = descargar_reporte_mantenimientos(
                    desde_i.strftime("%d/%m/%Y"), hasta_i.strftime("%d/%m/%Y"), tipo_mantenimiento_code,
                    intento=intento,
                )
                candidato = _parsear_excel_coes(raw_bytes)
                if _validar_rango_chunk(candidato, desde_i, hasta_i):
                    df_parte = candidato
                    break
                ultimo_error = "COES devolvió un reporte que no corresponde a este rango (posible dato desactualizado)"
                time.sleep(2 * (intento + 1))
            except Exception as e:
                ultimo_error = str(e)
                time.sleep(2 * (intento + 1))

        if df_parte is not None:
            partes.append(df_parte)
        else:
            errores.append(f"{desde_i.strftime('%d/%m/%Y')} - {hasta_i.strftime('%d/%m/%Y')}: {ultimo_error}")

        if progreso is not None:
            progreso.progress(i / len(rangos))

    if not partes:
        return pd.DataFrame(), errores
    df = pd.concat(partes, ignore_index=True, sort=False)
    return df.drop_duplicates().reset_index(drop=True), errores




def detectar_fila_encabezado(raw_bytes: bytes, max_filas_prueba: int = 10) -> int:
    """El reporte de COES trae unas filas de metadata (título, fecha inicial,
    fecha final) antes de la fila real de encabezados. La cantidad de esas
    filas puede variar según el estado del reporte (EJECUTADOS, PROGRAMADO
    DIARIO/SEMANAL/MENSUAL), así que en vez de asumir siempre "fila 4"
    buscamos la fila que más palabras clave de encabezado contiene.
    Devuelve el índice de fila (0-indexado) para usar como header= en
    pd.read_excel."""
    vista = pd.read_excel(io.BytesIO(raw_bytes), header=None, nrows=max_filas_prueba)
    palabras_clave = ("MANTENIMIENTO", "EMPRESA", "EQUIPO", "INICIO", "FIN", "UBICACION", "UBICACIÓN")
    mejor_fila, mejor_score = 0, -1
    for i in range(len(vista)):
        valores = [str(v).upper().strip() for v in vista.iloc[i] if pd.notna(v)]
        score = sum(1 for v in valores if any(k in v for k in palabras_clave))
        if score > mejor_score:
            mejor_score, mejor_fila = score, i
    return mejor_fila


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

    st.header("⏱️ Duración")
    filtrar_duracion = st.checkbox("Solo mantenimientos mayores a X días", value=False)
    umbral_dias = st.number_input(
        "Duración mayor a (días)", min_value=0, value=3, step=1, disabled=not filtrar_duracion
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
        fecha_desde,
        fecha_hasta,
        MANTENIMIENTO_ESTADOS[mantenimiento_sel],
        mantenimiento_sel,
    )

if "coes_fetch_params" in st.session_state:
    fecha_desde_busq, fecha_hasta_busq, code, mantenimiento_label = st.session_state["coes_fetch_params"]
    rangos_anuales = _generar_rangos_anuales(fecha_desde_busq, fecha_hasta_busq)

    with st.spinner(
        f"Descargando reporte de COES ({fecha_desde_busq.strftime('%d/%m/%Y')} a "
        f"{fecha_hasta_busq.strftime('%d/%m/%Y')}, {mantenimiento_label})"
        + (f" en {len(rangos_anuales)} tramos anuales..." if len(rangos_anuales) > 1 else "...")
    ):
        progreso = st.progress(0.0) if len(rangos_anuales) > 1 else None
        estado_texto = st.empty() if len(rangos_anuales) > 1 else None
        try:
            # El rango se pide año por año y se combina, porque COES trunca
            # silenciosamente los reportes que abarcan un rango de fechas muy
            # amplio (ver _generar_rangos_anuales / descargar_y_parsear_reporte).
            # Si algún año puntual falla (timeout, error del servidor, etc.), se
            # sigue con los demás en vez de perder toda la descarga.
            df, errores_descarga = descargar_y_parsear_reporte(
                fecha_desde_busq, fecha_hasta_busq, code, progreso=progreso, estado_texto=estado_texto
            )
        except Exception as e:
            st.error(f"❌ Error al descargar el reporte: {e}")
            st.stop()
        finally:
            if progreso is not None:
                progreso.empty()
            if estado_texto is not None:
                estado_texto.empty()

    if df.empty:
        st.error(
            "❌ El archivo descargado no se pudo leer como Excel, o vino vacío. "
            "Puede que COES haya devuelto una página de error en vez del reporte."
        )
        st.stop()

    if errores_descarga:
        st.warning(
            "⚠️ Algunos tramos anuales no se pudieron descargar, así que el reporte "
            "está incompleto para esos años:\n\n"
            + "\n".join(f"- {e}" for e in errores_descarga)
        )

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

    col_fecha_ini = (
        detectar_columna(df, ["FECHA", "INICIO"])
        or detectar_columna(df, ["INICIO"])
        or detectar_columna(df, ["FEC.", "INICIO"])
    )
    col_fecha_fin = (
        detectar_columna(df, ["FECHA", "FIN"])
        or detectar_columna(df, ["FIN"])
        or detectar_columna(df, ["FEC.", "FIN"])
    )

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

    c6, c7 = st.columns(2)
    with c6:
        col_fecha_ini = st.selectbox("Col. Fecha Inicio", options=[None] + list(df.columns),
                                      index=(list(df.columns).index(col_fecha_ini) + 1) if col_fecha_ini in df.columns else 0)
    with c7:
        col_fecha_fin = st.selectbox("Col. Fecha Fin", options=[None] + list(df.columns),
                                      index=(list(df.columns).index(col_fecha_fin) + 1) if col_fecha_fin in df.columns else 0)

    # --- Calcular duración en días (si hay columnas de fecha inicio/fin) ---
    col_duracion = None
    if col_fecha_ini and col_fecha_fin:
        f_ini = pd.to_datetime(df[col_fecha_ini], errors="coerce", dayfirst=True)
        f_fin = pd.to_datetime(df[col_fecha_fin], errors="coerce", dayfirst=True)
        df["Duración (días)"] = (f_fin - f_ini).dt.total_seconds() / 86400
        df["Duración (días)"] = df["Duración (días)"].round(2)
        col_duracion = "Duración (días)"
    else:
        st.warning(
            "⚠️ No se detectaron columnas de Fecha Inicio y Fecha Fin, así que no se puede calcular la "
            "duración. Selecciónalas manualmente arriba (Col. Fecha Inicio / Col. Fecha Fin)."
        )

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

    if filtrar_duracion and col_duracion:
        df_filtrado = df_filtrado[df_filtrado[col_duracion] > umbral_dias]

    df_filtrado = df_filtrado.reset_index(drop=True)

    # --- KPIs de duración ---
    if col_duracion:
        total_mayor_3 = (df[col_duracion] > 3).sum()
        k1, k2, k3 = st.columns(3)
        k1.metric("Total mantenimientos (reporte completo)", f"{len(df):,}")
        k2.metric("Mayores a 3 días (reporte completo)", f"{total_mayor_3:,}")
        k3.metric("En el resultado filtrado actual", f"{len(df_filtrado):,}")

    st.subheader(f"📋 Resultado filtrado: {len(df_filtrado):,} de {len(df):,} filas")

    if col_duracion:
        orden_por_duracion = st.checkbox("Ordenar por Duración (días), de mayor a menor", value=filtrar_duracion)
        if orden_por_duracion:
            df_filtrado = df_filtrado.sort_values(col_duracion, ascending=False).reset_index(drop=True)

        def resaltar_mayor_3(row):
            if pd.notna(row.get(col_duracion)) and row[col_duracion] > 3:
                return ["background-color: #ffe1e1"] * len(row)
            return [""] * len(row)

        # pandas.Styler tiene un límite duro de celdas que puede renderizar
        # (pd.options.styler.render.max_elements, ~262,144 por defecto). Con
        # varios años de datos combinados, el resultado filtrado puede superar
        # ese límite fácilmente y styler.apply() revienta con una excepción.
        # Por eso solo se usa el resaltado de color cuando el tamaño es
        # razonable; si no, se muestra la tabla sin colorear (los datos siguen
        # completos, solo se pierde el resaltado visual).
        LIMITE_CELDAS_STYLER = 250_000
        if df_filtrado.size <= LIMITE_CELDAS_STYLER:
            st.dataframe(
                df_filtrado.style.apply(resaltar_mayor_3, axis=1),
                use_container_width=True,
            )
        else:
            st.caption(
                f"ℹ️ El resultado tiene {df_filtrado.size:,} celdas, por lo que se omite el "
                f"resaltado de color (límite: {LIMITE_CELDAS_STYLER:,}) para que la tabla no falle. "
                "Los datos están completos igual; puedes filtrar más para ver el resaltado, "
                "o descargar el Excel/CSV más abajo."
            )
            st.dataframe(df_filtrado, use_container_width=True)
    else:
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

    # =====================================================================
    # MANTENIMIENTOS TÍPICOS / RECURRENTES (se repiten cada año)
    # =====================================================================
    st.divider()
    st.subheader("📆 Mantenimientos típicos: los que se repiten cada año")
    st.caption(
        "Agrupa por Central/Instalación (y opcionalmente por Tipo de Mantenimiento) + mes del año, "
        "y cuenta en cuántos años distintos del rango descargado aparece esa combinación. Si aparece "
        "en varios años seguidos, es un mantenimiento 'típico' recurrente y no un evento puntual."
    )

    anios_en_rango = fecha_hasta.year - fecha_desde.year + 1
    if anios_en_rango < 2:
        st.warning(
            f"⚠️ El rango de fechas que descargaste cubre solo {anios_en_rango} año(s). "
            f"Para detectar patrones anuales necesitas descargar varios años (por ejemplo 2022-2026) "
            f"en el panel izquierdo y volver a presionar 'Descargar y filtrar'."
        )

    col_central_auto = (
        detectar_columna(df, ["CENTRAL"])
        or detectar_columna(df, ["INSTALACION"])
        or detectar_columna(df, ["INSTALACIÓN"])
        or detectar_columna(df, ["NOMBRE", "EQUIPO"])
        or col_equipo
    )

    cA, cB, cC = st.columns(3)
    with cA:
        col_central = st.selectbox(
            "Columna que identifica la Central/Instalación específica",
            options=list(df.columns),
            index=list(df.columns).index(col_central_auto) if col_central_auto in df.columns else 0,
            help="Debe ser el nombre puntual de la central o instalación (ej. 'C.H. Machupicchu'), "
                 "no la categoría genérica de 'Tipo de Equipo'.",
        )
    with cB:
        min_anios = st.number_input(
            "Mínimo de años distintos para considerarlo 'típico'", min_value=2, value=2, step=1
        )
    with cC:
        agrupar_por_tipo_mantto = st.checkbox(
            "Agrupar también por Tipo de Mantenimiento", value=bool(col_mantto), disabled=(col_mantto is None)
        )

    if col_fecha_ini and col_central:
        f_ini_all = pd.to_datetime(df[col_fecha_ini], errors="coerce", dayfirst=True)
        df_pat = df.copy()
        df_pat["_Anio"] = f_ini_all.dt.year
        df_pat["_Mes"] = f_ini_all.dt.month
        df_pat["_MesNombre"] = df_pat["_Mes"].map(MESES_ES)
        df_pat = df_pat.dropna(subset=["_Anio", "_Mes"])

        group_cols = [col_central, "_Mes", "_MesNombre"]
        if agrupar_por_tipo_mantto and col_mantto:
            group_cols.insert(1, col_mantto)

        resumen = (
            df_pat.groupby(group_cols, dropna=False)["_Anio"]
            .agg(
                N_anios="nunique",
                Anios=lambda s: ", ".join(sorted(set(str(int(x)) for x in s))),
                N_registros="size",
            )
            .reset_index()
        )
        resumen["Típico (recurrente)"] = resumen["N_anios"] >= min_anios
        resumen = resumen.sort_values(["Típico (recurrente)", "N_anios"], ascending=[False, False]).reset_index(drop=True)
        resumen = resumen.rename(columns={
            "_Mes": "Mes (n°)", "_MesNombre": "Mes", "N_anios": "N° Años", "Anios": "Años en que ocurrió",
            "N_registros": "N° Registros",
        })

        solo_tipicos = st.checkbox("Mostrar solo los recurrentes/típicos", value=True)
        tabla_mostrar = resumen[resumen["Típico (recurrente)"]] if solo_tipicos else resumen

        st.dataframe(tabla_mostrar, use_container_width=True)
        st.caption(f"{len(tabla_mostrar):,} patrón(es) de {len(resumen):,} combinación(es) Central+Mes evaluadas.")

        buf_pat = io.BytesIO()
        resumen.to_excel(buf_pat, index=False, engine="openpyxl")
        st.download_button(
            "⬇️ Descargar patrón de recurrencia (.xlsx)",
            data=buf_pat.getvalue(),
            file_name="coes_mantenimientos_tipicos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info(
            "Para detectar el patrón anual, verifica que las columnas 'Col. Fecha Inicio' (arriba) y "
            "'Columna que identifica la Central' estén correctamente seleccionadas."
        )
else:
    st.info("Configura los filtros en el panel izquierdo y presiona **Descargar y filtrar**.")

if "coes_fetch_params" in st.session_state:
    if st.sidebar.button("🗑️ Limpiar resultado / nueva búsqueda"):
        del st.session_state["coes_fetch_params"]
        st.rerun()
