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
"""
)

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}

DEFAULT_OUTPUT = r"C:\Users\GZ6710\OneDrive - ENGIE\Escritorio\ENGIE\2026\Plexos\SCRIPT\PMENSUAL_2026.xlsx"

# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.header("⚙️ Configuración")
    output_path = st.text_input("Ruta de salida (.xlsx)", value=DEFAULT_OUTPUT)
    max_workers = st.slider("Descargas en paralelo", min_value=1, max_value=6, value=3)
    st.markdown("---")
    sheet_name = st.text_input("Nombre de hoja", value="MANTTOS")
    header_row = st.number_input(
        "Fila de encabezado (numeración de Excel, 1-indexed)", min_value=1, value=9
    )
    col_start = st.text_input("Columna inicial", value="B")
    col_end = st.text_input("Columna final", value="Q")

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
# HELPERS
# =========================
def generate_url(year, month):
    """
    Construye la URL de descarga del Programa Mensual para un año y mes dados.
    Ej: 2026, 7 -> .../Programa Mensual/2026/07_JULIO/Final/PMENSUAL_JULIO_2026.zip
    """
    mes_nombre = MESES[month]
    carpeta_mes = f"{month:02d}_{mes_nombre}"
    archivo_zip = f"PMENSUAL_{mes_nombre}_{year}.zip"
    url = (
        f"https://www.coes.org.pe/portal/browser/download?url="
        f"Operaci%C3%B3n%2FPrograma%20de%20Mantenimiento%2FPrograma%20Mensual%2F{year}%2F"
        f"{carpeta_mes}%2FFinal%2F{archivo_zip}"
    )
    return url


def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def download_and_extract(year, month, sheet_name, header_row, col_start, col_end):
    url = generate_url(year, month)
    mes_nombre = MESES[month]
    headers = {"User-Agent": "Mozilla/5.0"}

    response = requests.get(url, headers=headers, timeout=60)
    if response.status_code != 200:
        raise RuntimeError(f"HTTP {response.status_code} al descargar {mes_nombre} {year} (url: {url})")

    with zipfile.ZipFile(io.BytesIO(response.content)) as z:
        namelist = z.namelist()
        excel_files = [
            f for f in namelist
            if "INTERVENCIONES_(AGENTES)" in f.upper().replace(" ", "_")
            and f.endswith((".xlsx", ".xlsm", ".xls"))
        ]
        if not excel_files:
            raise FileNotFoundError(
                f"No se encontró Anexo1_Intervenciones_(Agentes) en el ZIP de {mes_nombre} {year}. "
                f"Archivos disponibles: {namelist}"
            )

        excel_name = excel_files[0]
        with z.open(excel_name) as excel_file:
            start_idx = col_letter_to_index(col_start)
            end_idx = col_letter_to_index(col_end)
            df = pd.read_excel(
                io.BytesIO(excel_file.read()),
                sheet_name=sheet_name,
                header=header_row - 1,  # 1-indexed en la UI -> 0-indexed para pandas
                usecols=list(range(start_idx, end_idx + 1)),
            )
            df = df.dropna(how="all").reset_index(drop=True)
            df["_year"] = year
            df["_month"] = month
            df["_periodo"] = f"{mes_nombre} {year}"
            return df


# =========================
# BOTÓN DE PROCESAMIENTO
# =========================
if st.button("🚀 Descargar y consolidar", type="primary"):
    if not periodos:
        st.error("Selecciona al menos un año y un mes.")
    else:
        results = []
        errors = []

        progress = st.progress(0.0)
        status = st.empty()
        t0 = time.time()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(
                    download_and_extract, y, m, sheet_name, header_row, col_start, col_end
                ): (y, m)
                for y, m in periodos
            }

            done = 0
            for future in as_completed(futures):
                y, m = futures[future]
                try:
                    results.append(future.result())
                except Exception as e:
                    errors.append(f"{MESES[m]} {y}: {e}")
                done += 1
                progress.progress(done / len(futures))
                status.text(f"Procesados {done}/{len(futures)} — último: {MESES[m]} {y}")

        elapsed = time.time() - t0
        status.empty()
        progress.empty()
        st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(periodos)} periodo(s), {max_workers} en paralelo)")

        if errors:
            with st.expander("⚠️ Errores", expanded=True):
                for e in errors:
                    st.text(e)

        if results:
            df_final = pd.concat(results, ignore_index=True)
            st.success(f"✅ {len(df_final):,} filas consolidadas de {len(results)} periodo(s).")
            st.dataframe(df_final.head(300), use_container_width=True)

            # Guarda directamente en la ruta local (sobrescribe, igual que el script original)
            try:
                df_final.to_excel(output_path, index=False)
                st.success(f"💾 Guardado en: {output_path}")
            except Exception as e:
                st.warning(f"No se pudo guardar en la ruta local ({e}). Usa el botón de descarga.")

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
