# -*- coding: utf-8 -*-
import io
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import pandas as pd
import streamlit as st

# =========================
# CONFIG DE PÁGINA
# =========================
st.set_page_config(page_title="COES - Programa Mensual (Intervenciones)", layout="wide")
st.title("🛠️ Descarga y Consolidación — Programa Mensual COES (Intervenciones)")

st.markdown(
    """
Descarga automáticamente el archivo **Anexo1_Intervenciones_(Agentes)** del Programa
Mensual de Mantenimiento publicado por COES, para uno o varios periodos (año/mes),
y consolida todo en una sola tabla.

Como el nombre exacto del ZIP puede variar de un mes a otro (mayúsculas, tildes, etc.),
la app prueba automáticamente varias variantes del nombre antes de descargar.
"""
)

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

# Nombres de mes alternativos que a veces usa COES para el mismo mes
MESES_ALT = {
    9: ["SEPTIEMBRE"],
}

# =========================
# CONFIG FIJA (antes eran inputs en el sidebar; ahora se fijan en el código)
# =========================
SHEET_NAME = "MANTTOS"
HEADER_ROW = 9      # numeración de Excel, 1-indexed
COL_START = "B"
COL_END = "Q"

# =========================
# SIDEBAR (simplificado)
# =========================
with st.sidebar:
    st.header("⚙️ Configuración")
    max_workers = st.slider("Descargas en paralelo", min_value=1, max_value=6, value=3)
    show_debug = st.checkbox("Mostrar variantes de URL probadas (debug)", value=False)

# =========================
# SELECCIÓN DE PERIODOS
# =========================
st.subheader("📅 Periodos a descargar")
col1, col2 = st.columns(2)
with col1:
    years = st.multiselect("Año(s)", options=list(range(2022, 2031)), default=[2026])
with col2:
    months = st.multiselect(
        "Mes(es)",
        options=list(MESES.keys()),
        default=[7],
        format_func=lambda m: MESES[m],
    )

periodos = [(y, m) for y in sorted(years) for m in sorted(months)]
st.caption(f"Se descargarán {len(periodos)} periodo(s).")


# =========================
# HELPERS: generación de candidatos de URL
# =========================
def get_month_name_variants(month: int):
    """Todas las variantes de nombre de mes a probar para un mes dado."""
    return [MESES[month]] + MESES_ALT.get(month, [])


def generate_url_candidates(year: int, month: int):
    """
    Genera una lista de URLs candidatas para el ZIP del Programa Mensual,
    probando variantes de nombre de mes y de mayúsculas/minúsculas del archivo.
    """
    candidates = []
    for mes_nombre in get_month_name_variants(month):
        carpeta_mes = f"{month:02d}_{mes_nombre}"
        for case_variant in {mes_nombre, mes_nombre.title(), mes_nombre.capitalize()}:
            archivo_zip = f"PMENSUAL_{case_variant}_{year}.zip"
            url = (
                f"https://www.coes.org.pe/portal/browser/download?url="
                f"Operaci%C3%B3n%2FPrograma%20de%20Mantenimiento%2FPrograma%20Mensual%2F{year}%2F"
                f"{carpeta_mes}%2FFinal%2F{archivo_zip}"
            )
            candidates.append(url)

    # dedupe preservando el orden
    seen = set()
    unique = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            unique.append(c)
    return unique


def find_working_url(year: int, month: int, headers: dict, timeout: int = 20):
    """
    Prueba cada URL candidata con un HEAD (o GET en streaming si el servidor
    no soporta HEAD) y devuelve la primera que responda 200, junto con el
    detalle completo de lo probado (para diagnóstico).
    """
    tried = []
    for url in generate_url_candidates(year, month):
        status = None
        try:
            resp = requests.head(url, headers=headers, timeout=timeout, allow_redirects=True)
            status = resp.status_code
            if status in (405, 501):  # servidor no soporta HEAD -> probar GET en streaming
                resp = requests.get(url, headers=headers, timeout=timeout, stream=True)
                status = resp.status_code
                resp.close()
        except Exception as e:
            tried.append((url, f"ERROR: {e}"))
            continue

        tried.append((url, status))
        if status == 200:
            return url, tried

    return None, tried


def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


# =========================
# DESCARGA Y PROCESAMIENTO
# =========================
def download_and_extract(year, month):
    mes_display = MESES[month]
    headers = {"User-Agent": "Mozilla/5.0"}

    url, tried = find_working_url(year, month, headers)
    if url is None:
        detail = "\n".join(f"  [{status}] {u}" for u, status in tried)
        raise RuntimeError(
            f"No se encontró un ZIP válido para {mes_display} {year} tras probar "
            f"{len(tried)} variante(s):\n{detail}"
        )

    response = requests.get(url, headers=headers, timeout=60)
    if response.status_code != 200:
        raise RuntimeError(f"HTTP {response.status_code} al descargar {mes_display} {year} (url: {url})")

    with zipfile.ZipFile(io.BytesIO(response.content)) as z:
        namelist = z.namelist()
        excel_files = [
            f for f in namelist
            if "INTERVENCIONES_(AGENTES)" in f.upper().replace(" ", "_")
            and f.endswith((".xlsx", ".xlsm", ".xls"))
        ]
        if not excel_files:
            raise FileNotFoundError(
                f"No se encontró Anexo1_Intervenciones_(Agentes) en el ZIP de {mes_display} {year} "
                f"(url usada: {url}). Archivos disponibles: {namelist}"
            )

        excel_name = excel_files[0]
        with z.open(excel_name) as excel_file:
            start_idx = col_letter_to_index(COL_START)
            end_idx = col_letter_to_index(COL_END)
            df = pd.read_excel(
                io.BytesIO(excel_file.read()),
                sheet_name=SHEET_NAME,
                header=HEADER_ROW - 1,  # 1-indexed -> 0-indexed para pandas
                usecols=list(range(start_idx, end_idx + 1)),
            )
            df = df.dropna(how="all").reset_index(drop=True)
            df["_year"] = year
            df["_month"] = month
            df["_periodo"] = f"{mes_display} {year}"

    return df, url, tried


# =========================
# BOTÓN DE PROCESAMIENTO
# =========================
if st.button("🚀 Descargar y consolidar", type="primary"):
    if not periodos:
        st.error("Selecciona al menos un año y un mes.")
    else:
        results = []
        errors = []
        debug_info = []

        progress = st.progress(0.0)
        status = st.empty()
        t0 = time.time()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(download_and_extract, y, m): (y, m)
                for y, m in periodos
            }

            done = 0
            for future in as_completed(futures):
                y, m = futures[future]
                periodo_label = f"{MESES[m]} {y}"
                try:
                    df, used_url, tried = future.result()
                    results.append(df)
                    debug_info.append({"periodo": periodo_label, "used_url": used_url, "tried": tried})
                except Exception as e:
                    errors.append(f"{periodo_label}: {e}")
                done += 1
                progress.progress(done / len(futures))
                status.text(f"Procesados {done}/{len(futures)} — último: {periodo_label}")

        elapsed = time.time() - t0
        status.empty()
        progress.empty()
        st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(periodos)} periodo(s), {max_workers} en paralelo)")

        if errors:
            with st.expander("⚠️ Errores", expanded=True):
                for e in errors:
                    st.text(e)

        if show_debug and debug_info:
            with st.expander("🔎 Variantes de URL probadas por periodo", expanded=False):
                for info in debug_info:
                    st.markdown(f"**{info['periodo']}** — usada: `{info['used_url']}`")
                    for u, s in info["tried"]:
                        st.text(f"  [{s}] {u}")

        if results:
            df_final = pd.concat(results, ignore_index=True)
            st.success(f"✅ {len(df_final):,} filas consolidadas de {len(results)} periodo(s).")
            st.dataframe(df_final.head(300), use_container_width=True)

            # Descarga xlsx
            buf_xlsx = io.BytesIO()
            df_final.to_excel(buf_xlsx, index=False, engine="openpyxl")
            st.download_button(
                "⬇️ Descargar PMENSUAL.xlsx",
                data=buf_xlsx.getvalue(),
                file_name="PMENSUAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Descarga csv
            buf_csv = df_final.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Descargar PMENSUAL.csv",
                data=buf_csv,
                file_name="PMENSUAL.csv",
                mime="text/csv",
            )
        else:
            st.error("❌ No se consolidó ningún periodo.")
else:
    st.info("Selecciona los periodos y presiona **Descargar y consolidar**.")
