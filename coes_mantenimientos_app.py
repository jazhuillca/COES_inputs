# -*- coding: utf-8 -*-
"""
Streamlit — Descarga de "Medidores de Generación" (COES)

Ejecutar:
    pip install streamlit requests pandas openpyxl
    streamlit run app_medidores_generacion.py
"""

import io
import re
import time
import unicodedata
from calendar import monthrange
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

# ─────────────────────────────────────────────────────────────
#  CONSTANTES DEL ENDPOINT / FORMULARIO (extraídas del HTML real)
# ─────────────────────────────────────────────────────────────
TIPOS_GENERACION = {
    "EÓLICA": "4",
    "HIDROELÉCTRICA": "1",
    "SOLAR": "3",
    "TERMOELÉCTRICA": "2",
}

CENTRAL_OPCIONES = {
    "TODOS": "0",
    "COES": "1",
    "GENERACION RER": "3",
}

# Mapeo nombre -> ID interno COES, extraído del <select id="cbEmpresas"> real del formulario
EMPRESAS = {
    "ACCIONA ENERGIA PERU S.A.C": "14260",
    "ADINELSA ADN": "69",
    "AGRO INDUSTRIAL PARAMONGA S.A.": "15214",
    "AGROAURORA S.A.C.": "11772",
    "AGROINDUSTRIAS SAN JACINTO S.A.A.": "12439",
    "AGROLMOS SOCIEDAD ANONIMA - AGROLMOS S.A.": "11777",
    "AGUAS Y ENERGIA PERU": "10481",
    "ANDEAN POWER S.A.C.": "12056",
    "ASOCIACIÓN SANTA LUCIA DE CHACAS": "12362",
    "ATRIA ENERGIA S.A.C.": "13196",
    "BIOENERGIA DEL CHIRA S.A.": "12896",
    "CASA GRANDE S.A.A.": "10628",
    "CELARIS ENERGY S.A.": "15707",
    "CELEPSA": "10420",
    "CELEPSA RENOVABLES S.R.L.": "12708",
    "CENTRALES SANTA ROSA S.A.C.": "13165",
    "CHINANGO S.A.C.": "10901",
    "COENERGY S.A.C.": "15571",
    "COLCA SOLAR S.A.C.": "12584",
    "CORPORACION MINERA DEL PERU S.A.": "11095",
    "EGASA": "17",
    "EGEJUNIN": "11153",
    "EGEMSA": "58",
    "EGESUR": "19",
    "ELECTRICA YANAPAMPA SAC": "10684",
    "ELECTRO ORIENTE": "30",
    "ELECTRO SUR ESTE": "27",
    "ELECTRO UCAYALI": "40",
    "ELECTRO ZAÑA S.A.C.": "11228",
    "ELECTRONOROESTE S.A.": "23",
    "ELECTRONORTE S.A.": "24",
    "ELECTROPERU": "2",
    "EMPRESA DE GENERACION ELECTRICA CANCHAYLLO SAC": "11389",
    "EMPRESA DE GENERACION ELECTRICA SANTA ANA S.A.C.": "14173",
    "EMPRESA DE GENERACION HUALLAGA": "11412",
    "EMPRESA DE GENERACION HUANZA": "206",
    "EMPRESA ELECTRICA AGUA AZUL": "11544",
    "EMPRESA ELECTRICA RIO DOBLE": "11058",
    "ENEL GENERACION PIURA S.A.": "12097",
    "ENERGÍA EÓLICA S.A.": "10552",
    "ENERGIA RENOVABLE DEL SUR S.A.": "11429",
    "ENERGIA RENOVABLE LA JOYA S.A.": "12884",
    "ENGIE ENERGIA PERU S.A.A.": "15624",
    "EOLICA CARAVELI S.A.C.": "15392",
    "FENIX POWER PERÚ": "10725",
    "GENERACIÓN ANDINA S.A.C.": "11527",
    "GENERADORA DE ENERGÍA DEL PERÚ": "10647",
    "GM OPERACIONES S.A.C.": "14973",
    "GR CORTARRAMA SOCIEDAD ANONIMA CERRADA": "11981",
    "GR PAINO SOCIEDAD ANONIMA CERRADA": "11840",
    "GR TARUCA SOCIEDAD ANONIMA CERRADA": "11841",
    "GR VALE S.A.C.": "12974",
    "HIDROCAÑETE S.A.": "10974",
    "HIDROELECTRICA HUANCHOR S.A.C.": "11258",
    "HIDROELECTRICA RAPAZ S.A.C.": "11644",
    "HUAURA POWER GROUP S.A.": "11444",
    "HYDRO GLOBAL PERÚ S.A.C.": "11940",
    "HYDRO PATAPO S.A.C.": "12364",
    "ILLAPU ENERGY": "11149",
    "INFRAESTRUCTURAS Y ENERGIAS DEL PERU S.A.C.": "11528",
    "INLAND ENERGY SAC": "12634",
    "INTI JOYA S.A.C.": "15849",
    "INVERSIONES SHAQSHA S.A.C.": "12624",
    "JOYA SOLAR S.A.C.": "13726",
    "KALLPA GENERACION S.A.": "12479",
    "KONDU SAC": "14807",
    "LA VIRGEN": "11185",
    "LUZ DEL SUR": "13",
    "MAJA ENERGIA S.A.C.": "10916",
    "MAJES ARCUS S.A.C.": "13965",
    "MINERA CERRO VERDE": "67",
    "MINERA CORONA": "108",
    "MOQUEGUA FV S.A.C.": "11217",
    "ORAZUL ENERGY PERÚ": "12480",
    "ORYGEN PERU S.A.A.": "15259",
    "PANAMERICANA SOLAR SAC.": "11102",
    "PARQUE EOLICO MARCONA S.A.C.": "13120",
    "PARQUE EOLICO TRES HERMANAS S.A.C.": "11218",
    "PETRAMAS": "11063",
    "PETROPERU": "149",
    "PLANTA DE RESERVA FRIA DE GENERACION DE ETEN S.A.": "11323",
    "REPARTICIÓN ARCUS S.A.C.": "13966",
    "SAMAY I S.A.C": "15009",
    "SAN GABAN": "61",
    "SDE PIURA": "10913",
    "SDF ENERGIA SAC": "10587",
    "SHOUGESA": "8",
    "SINERSA": "138",
    "STATKRAFT S.A": "12758",
    "TACNA SOLAR SAC.": "11103",
    "TERMOCHILCA S.A.C.": "15080",
    "TERMOSELVA": "10",
    "TRANSMISION ANDINA DE GENERACION S.A.C.": "15683",
    "UNACEM PERU S.A.": "14342",
    "VARI ENERGIA S.A.C.": "13984",
    "YURA": "167",
    "AGRO INDUSTRIAL PARAMONGA": "10422",
    "CEMENTO ANDINO": "180",
    "CERRO DEL AGUILA S.A.": "11146",
    "COGENERACION OQUENDO SAC": "15014",
    "EDEGEL": "4",
    "EEPSA": "5",
    "EGENOR": "9",
    "ELECTRICA SANTA ROSA": "76",
    "EMPRESA CONCESIONARIA ENERGIA LIMPIA SAC": "11563",
    "EMPRESA DE GENERACIÓN ELÉCTRICA CHEVES S.A.": "10636",
    "EMPRESA DE GENERACION ELECTRICA RIO BAÑOS S.A.C.": "11129",
    "EMPRESA DE GENERACION ELECTRICA SANTA ANA": "11509",
    "ENEL GENERACION PERU S.A.A.": "12096",
    "ENEL GREEN POWER PERU S.A.": "11395",
    "ENEL GREEN POWER PERU S.A.C": "13783",
    "ENERSUR": "18",
    "ENGIE": "48",
    "HIDROELECTRICA SANTA CRUZ": "10582",
    "HIDROMARAÑON": "11064",
    "KALLPA GENERACION": "47",
    "MAJES ARCUS": "11100",
    "MAPLE ETANOL": "10755",
    "ORAZUL ENERGY": "12190",
    "PARQUE EOLICO MARCONA S.R.L.": "11053",
    "PERUANA DE INVERSIONES EN ENERGIAS RENOVABLES S.A.": "10984",
    "REPARTICIÓN ARCUS": "11101",
    "SAMAY I S.A.": "11486",
    "SN POWER": "6",
    "STATKRAFT": "11567",
    "TERMOCHILCA": "10767",
    "UNION ANDINA DE CEMENTO": "11894",
}

PARAMETROS_EXPORTAR = {
    "Potencia Activa (MW)": "1",
    "Potencia Reactiva (MW)": "5",
    "Servicios Auxiliares": "3",
    "Potencia Reactiva Capacitiva (MVAR)": "2",
    "Potencia Reactiva Inductiva (MVAR)": "4",
}

FORMATOS = {
    "Excel Horizontal": "1",
    "Excel Vertical": "2",
    "CSV": "3",
}

FORMATO_EXTENSION = {"1": "xlsx", "2": "xlsx", "3": "csv"}

URL_PAGINA          = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion"
URL_VALIDAR_EXPORT  = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/validarexportacion"
URL_EXPORTAR        = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/exportar"
URL_DESCARGAR       = "https://www.coes.org.pe/Portal/mediciones/medidoresgeneracion/descargar"

MENSAJES_VALIDACION = {
    2: "El lapso de tiempo no puede ser mayor a 1 mes.",
    3: "Para la exportación a CSV solo debe seleccionar un parámetro.",
    4: "Seleccione un parámetro a exportar.",
    -1: "Ha ocurrido un error en el servidor al validar la exportación.",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": URL_PAGINA,
    "Origin": "https://www.coes.org.pe",
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "application/json, text/javascript, */*; q=0.01",
}


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    s.get(URL_PAGINA, timeout=30)  # cookies de sesión
    return s


def _nombre_desde_content_disposition(resp: requests.Response, default: str) -> str:
    cd = resp.headers.get("Content-Disposition", "")
    m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd)
    if m:
        return m.group(1)
    return default


def descargar_medidores_generacion(
    fecha_inicial: str,
    fecha_final: str,
    empresas: str,
    tipos_generacion: str,
    central: str,
    parametros: str,
    tipo: str,
    tipos_empresa: str = "",
):
    """
    Replica el flujo real del sitio (medidores.js -> exportarFormato):
      1) POST validarexportacion
      2) POST exportar
      3) GET descargar?tipo=...

    Devuelve (contenido_bytes, nombre_archivo, mensaje_error).
    """
    s = _session()

    # ── Paso 1: validar ─────────────────────────────────────────
    payload_validar = {
        "formato": tipo,
        "fechaInicial": fecha_inicial,
        "fechaFinal": fecha_final,
        "parametros": parametros,
    }
    try:
        resp1 = s.post(URL_VALIDAR_EXPORT, data=payload_validar, timeout=60)
    except requests.RequestException as e:
        return None, None, f"Error de conexión en validarexportacion: {e}"

    if resp1.status_code != 200:
        return None, None, f"HTTP {resp1.status_code} en validarexportacion: {resp1.text[:300]}"

    try:
        resultado_validar = resp1.json()
    except ValueError:
        return None, None, f"Respuesta inesperada en validarexportacion: {resp1.text[:300]}"

    if resultado_validar != 1:
        mensaje = MENSAJES_VALIDACION.get(
            resultado_validar, f"Validación falló con código {resultado_validar!r}"
        )
        return None, None, mensaje

    # ── Paso 2: exportar (genera el archivo en la sesión del servidor) ──
    payload_exportar = {
        "fechaInicial": fecha_inicial,
        "fechaFinal": fecha_final,
        "tiposEmpresa": tipos_empresa,
        "empresas": empresas,
        "tiposGeneracion": tipos_generacion,
        "central": central,
        "parametros": parametros,
        "tipo": tipo,
    }
    try:
        resp2 = s.post(URL_EXPORTAR, data=payload_exportar, timeout=90)
    except requests.RequestException as e:
        return None, None, f"Error de conexión en exportar: {e}"

    if resp2.status_code != 200:
        return None, None, f"HTTP {resp2.status_code} en exportar: {resp2.text[:300]}"

    try:
        resultado_exportar = resp2.json()
    except ValueError:
        return None, None, f"Respuesta inesperada en exportar: {resp2.text[:300]}"

    if str(resultado_exportar) != "1":
        return None, None, f"exportar devolvió un error: {resultado_exportar!r}"

    # ── Paso 3: descargar el archivo ya generado ────────────────
    try:
        resp3 = s.get(URL_DESCARGAR, params={"tipo": tipo}, timeout=90)
    except requests.RequestException as e:
        return None, None, f"Error de conexión en descargar: {e}"

    if resp3.status_code != 200:
        return None, None, f"HTTP {resp3.status_code} en descargar: {resp3.text[:300]}"

    ext_default = FORMATO_EXTENSION.get(tipo, "xlsx")
    nombre_default = (
        f"MedidoresGeneracion_{fecha_inicial.replace('/', '')}_"
        f"{fecha_final.replace('/', '')}.{ext_default}"
    )
    nombre = _nombre_desde_content_disposition(resp3, default=nombre_default)
    nombre = unicodedata.normalize("NFKD", nombre).encode("ascii", "ignore").decode()

    return resp3.content, nombre, None


@st.cache_data(show_spinner=False, ttl=3600)
def _descargar_tramo_cacheado(
    fecha_inicial: str,
    fecha_final: str,
    empresas: str,
    tipos_generacion: str,
    central: str,
    parametros: str,
    tipo: str,
):
    """Wrapper cacheado de descargar_medidores_generacion: mismos params ->
    no se vuelve a golpear el servidor de COES (clave para que los botones
    de descarga, que fuerzan un rerun de todo el script, no disparen
    peticiones nuevas)."""
    return descargar_medidores_generacion(
        fecha_inicial=fecha_inicial,
        fecha_final=fecha_final,
        empresas=empresas,
        tipos_generacion=tipos_generacion,
        central=central,
        parametros=parametros,
        tipo=tipo,
    )


def dividir_en_meses(fecha_inicio: date, fecha_fin: date):
    """
    Divide [fecha_inicio, fecha_fin] en tramos que respetan el límite de
    1 mes que exige el servidor. Cada tramo va del día de inicio hasta el
    mismo día del mes siguiente menos uno (o el fin real, lo que sea antes).
    """
    tramos = []
    actual = fecha_inicio
    while actual <= fecha_fin:
        # Fin de tramo: mismo día un mes después, menos 1 día
        mes = actual.month + 1
        anio = actual.year
        if mes > 12:
            mes = 1
            anio += 1
        ultimo_dia_mes_destino = monthrange(anio, mes)[1]
        dia = min(actual.day, ultimo_dia_mes_destino)
        fin_tramo = date(anio, mes, dia) - timedelta(days=1)

        if fin_tramo > fecha_fin:
            fin_tramo = fecha_fin

        tramos.append((actual, fin_tramo))
        actual = fin_tramo + timedelta(days=1)

    return tramos


# ─────────────────────────────────────────────────────────────
#  INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Medidores de Generación — COES", page_icon="⚡", layout="centered")

st.title("⚡ Medidores de Generación — COES")
st.caption("Descarga directa desde el portal COES (mediciones/medidoresgeneracion)")

# IMPORTANTE: estos filtros van FUERA de un st.form a propósito.
# Dentro de un st.form, Streamlit no vuelve a ejecutar el script hasta
# que se presiona el botón de submit, así que un checkbox "TODOS" no
# podría habilitar/deshabilitar el multiselect en tiempo real. Al
# dejarlos sueltos, cada clic dispara un rerun inmediato.

st.subheader("Rango de fechas")
st.caption(
    "El portal solo permite consultar 1 mes por request; si eliges un rango "
    "mayor, la app descarga mes por mes y consolida todo en un solo Excel."
)
col1, col2 = st.columns(2)
with col1:
    fecha_inicial = st.date_input("Fecha inicial", value=date.today().replace(day=1))
with col2:
    fecha_final = st.date_input("Fecha final", value=date.today())

st.divider()
st.subheader(f"Empresa ({len(EMPRESAS)} agentes)")
todos_empresas = st.checkbox("TODOS los agentes", value=True, key="chk_empresas")
empresas_sel = st.multiselect(
    "Selecciona empresa(s)",
    options=sorted(EMPRESAS.keys()),
    default=[],
    disabled=todos_empresas,
    help="Escribe para filtrar por nombre.",
)

st.divider()
st.subheader("Tipo de Generación")
todos_tipo_gen = st.checkbox("TODOS los tipos", value=True, key="chk_tipo_gen")
tipo_gen_sel = st.multiselect(
    "Selecciona tipo(s) de generación",
    options=list(TIPOS_GENERACION.keys()),
    default=list(TIPOS_GENERACION.keys()),
    disabled=todos_tipo_gen,
)

st.divider()
st.subheader("Central / Alcance")
central_sel = st.radio(
    "Filtro de central",
    options=list(CENTRAL_OPCIONES.keys()),
    index=0,  # TODOS
    horizontal=True,
)

st.divider()
st.subheader("Parámetro (a exportar)")
todos_parametros = st.checkbox("TODOS los parámetros", value=True, key="chk_param")
parametros_sel = st.multiselect(
    "Selecciona parámetro(s)",
    options=list(PARAMETROS_EXPORTAR.keys()),
    default=list(PARAMETROS_EXPORTAR.keys()),
    disabled=todos_parametros,
)

st.divider()
st.subheader("Formato de salida")
formato_sel = st.radio(
    "Formato",
    options=list(FORMATOS.keys()),
    index=0,  # Excel Horizontal
    horizontal=True,
)

st.divider()
with st.expander("⚙️ Opciones avanzadas", expanded=False):
    max_workers = st.slider("Descargas en paralelo (para rangos >1 mes)", 1, 6, 3)

st.divider()
submitted = st.button("📥 Descargar", use_container_width=True, type="primary")


# ─────────────────────────────────────────────────────────────
#  IMPORTANTE: persistencia con session_state
#  Al presionar cualquier st.download_button, Streamlit vuelve a correr
#  todo el script desde arriba y `submitted` vuelve a False. Si el bloque
#  de resultados dependiera solo de `if submitted:`, desaparecería justo
#  antes de que el navegador termine de guardar el archivo. Por eso los
#  parámetros de la última búsqueda se guardan en session_state, y el
#  bloque de resultados se muestra mientras existan ahí — la descarga
#  real usa caché (_descargar_tramo_cacheado), así que no se vuelve a
#  golpear el servidor de COES en cada rerun.
# ─────────────────────────────────────────────────────────────
if submitted:
    if fecha_inicial > fecha_final:
        st.error("La fecha inicial no puede ser posterior a la fecha final.")
        st.stop()
    if not todos_empresas and not empresas_sel:
        st.error("Selecciona al menos una empresa (o marca TODOS los agentes).")
        st.stop()
    if not todos_tipo_gen and not tipo_gen_sel:
        st.error("Selecciona al menos un tipo de generación (o marca TODOS).")
        st.stop()
    if not todos_parametros and not parametros_sel:
        st.error("Selecciona al menos un parámetro (o marca TODOS).")
        st.stop()

    empresas_val = (
        ",".join(EMPRESAS.values()) if todos_empresas
        else ",".join(EMPRESAS[nombre] for nombre in empresas_sel)
    )
    tipo_gen_val = (
        ",".join(TIPOS_GENERACION.values())
        if todos_tipo_gen
        else ",".join(TIPOS_GENERACION[t] for t in tipo_gen_sel)
    )
    central_val = CENTRAL_OPCIONES[central_sel]
    parametros_val = (
        ",".join(PARAMETROS_EXPORTAR.values())
        if todos_parametros
        else ",".join(PARAMETROS_EXPORTAR[p] for p in parametros_sel)
    )
    formato_val = FORMATOS[formato_sel]

    st.session_state["medidores_params"] = dict(
        fecha_inicial=fecha_inicial,
        fecha_final=fecha_final,
        empresas_val=empresas_val,
        tipo_gen_val=tipo_gen_val,
        central_val=central_val,
        parametros_val=parametros_val,
        formato_val=formato_val,
        max_workers=max_workers,
    )


if "medidores_params" in st.session_state:
    p = st.session_state["medidores_params"]
    tramos = dividir_en_meses(p["fecha_inicial"], p["fecha_final"])

    if len(tramos) == 1:
        # Rango de 1 mes o menos: descarga directa
        with st.spinner("Consultando el portal COES..."):
            contenido, nombre, error = _descargar_tramo_cacheado(
                fecha_inicial=p["fecha_inicial"].strftime("%d/%m/%Y"),
                fecha_final=p["fecha_final"].strftime("%d/%m/%Y"),
                empresas=p["empresas_val"],
                tipos_generacion=p["tipo_gen_val"],
                central=p["central_val"],
                parametros=p["parametros_val"],
                tipo=p["formato_val"],
            )

        if error:
            st.error(f"No se pudo descargar: {error}")
        else:
            st.success(f"Listo — {nombre}")
            st.download_button(
                "⬇️ Guardar archivo",
                data=contenido,
                file_name=nombre,
                use_container_width=True,
            )

    else:
        # Rango mayor a 1 mes: el servidor solo acepta 1 mes por request,
        # así que se descarga tramo por tramo EN PARALELO y se consolida
        # todo en un solo archivo.
        st.info(
            f"El rango pedido abarca {len(tramos)} meses. El portal COES solo "
            f"permite descargar 1 mes a la vez, así que se descargarán en "
            f"paralelo ({p['max_workers']} a la vez) y se consolidará todo "
            f"en un solo archivo."
        )

        def _descargar_y_leer(ini, fin):
            rango_str = f"{ini.strftime('%d/%m/%Y')} - {fin.strftime('%d/%m/%Y')}"
            contenido, nombre, error = _descargar_tramo_cacheado(
                fecha_inicial=ini.strftime("%d/%m/%Y"),
                fecha_final=fin.strftime("%d/%m/%Y"),
                empresas=p["empresas_val"],
                tipos_generacion=p["tipo_gen_val"],
                central=p["central_val"],
                parametros=p["parametros_val"],
                tipo=p["formato_val"],
            )
            if error:
                raise RuntimeError(error)

            if p["formato_val"] == "3":  # CSV
                df = pd.read_csv(io.BytesIO(contenido), sep=None, engine="python")
            else:  # Excel Horizontal o Vertical
                df = pd.read_excel(io.BytesIO(contenido))
            df.insert(0, "Tramo_Consultado", rango_str)
            return df

        progreso = st.progress(0.0)
        estado = st.empty()
        t0 = time.time()

        dataframes = []
        errores = []  # (rango_str, mensaje)

        with ThreadPoolExecutor(max_workers=p["max_workers"]) as executor:
            futures = {
                executor.submit(_descargar_y_leer, ini, fin): (ini, fin)
                for ini, fin in tramos
            }

            done = 0
            for future in as_completed(futures):
                ini, fin = futures[future]
                rango_str = f"{ini.strftime('%d/%m/%Y')} - {fin.strftime('%d/%m/%Y')}"
                try:
                    dataframes.append(future.result())
                except Exception as e:
                    errores.append((rango_str, str(e)))
                done += 1
                progreso.progress(done / len(tramos))
                estado.text(f"Procesados {done}/{len(tramos)} — último: {rango_str}")

        elapsed = time.time() - t0
        estado.empty()
        progreso.empty()
        st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(tramos)} tramo(s), {p['max_workers']} en paralelo)")

        if errores:
            with st.expander(f"⚠️ {len(errores)} tramo(s) fallaron", expanded=True):
                for rango_str, msg in errores:
                    st.write(f"- **{rango_str}**: {msg}")

        if dataframes:
            df_final = pd.concat(dataframes, ignore_index=True)

            st.success(
                f"Listo — {len(dataframes)} de {len(tramos)} tramos consolidados "
                f"({len(df_final):,} filas en total)."
            )
            st.dataframe(df_final.head(300), use_container_width=True)

            nombre_base = (
                f"MedidoresGeneracion_{p['fecha_inicial'].strftime('%Y%m%d')}_"
                f"{p['fecha_final'].strftime('%Y%m%d')}_consolidado"
            )

            buf_xlsx = io.BytesIO()
            with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="MedidoresGeneracion")
            st.download_button(
                "⬇️ Descargar Excel consolidado",
                data=buf_xlsx.getvalue(),
                file_name=f"{nombre_base}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            buf_csv = df_final.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Descargar CSV consolidado",
                data=buf_csv,
                file_name=f"{nombre_base}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.error("No se pudo descargar ningún tramo.")

    st.divider()
    if st.button("🗑️ Limpiar resultado / nueva búsqueda"):
        del st.session_state["medidores_params"]
        st.rerun()
else:
    st.info("Configura los filtros arriba y presiona **📥 Descargar**.")
