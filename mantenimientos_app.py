# -*- coding: utf-8 -*-
import io
import time
import zipfile
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import pandas as pd
import streamlit as st

# =========================
# CONFIG DE PÁGINA
# =========================
st.set_page_config(page_title="COES - Mantenimientos", layout="wide")
st.title("🛠️ Programas de Mantenimiento COES")

MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}
MESES_ALT = {9: ["SEPTIEMBRE"]}

BASE_URL = "https://www.coes.org.pe/portal/browser/download?url="


def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


tab_mensual, tab_pai = st.tabs(["📅 Programa Mensual (Intervenciones)", "📆 Programa Anual (PAI)"])

# ======================================================================
# TAB 1: PROGRAMA MENSUAL (Anexo1_Intervenciones_(Agentes))
# ======================================================================
with tab_mensual:
    st.markdown(
        """
    Descarga el archivo **Anexo1_Intervenciones_(Agentes)** del Programa Mensual de
    Mantenimiento, para uno o varios periodos (año/mes), y consolida todo en una sola tabla.
    """
    )

    SHEET_MENSUAL = "MANTTOS"
    HEADER_MENSUAL = 9
    COL_START_MENSUAL, COL_END_MENSUAL = "B", "Q"

    with st.expander("⚙️ Opciones", expanded=False):
        max_workers_m = st.slider("Descargas en paralelo", 1, 6, 3, key="workers_mensual")
        show_debug_m = st.checkbox("Mostrar variantes de URL probadas (debug)", key="debug_mensual")

    col1, col2 = st.columns(2)
    with col1:
        years_m = st.multiselect("Año(s)", options=list(range(2022, 2031)), default=[2026], key="years_mensual")
    with col2:
        months_m = st.multiselect(
            "Mes(es)", options=list(MESES.keys()), default=[7],
            format_func=lambda m: MESES[m], key="months_mensual",
        )
    periodos_m = [(y, m) for y in sorted(years_m) for m in sorted(months_m)]
    st.caption(f"Se descargarán {len(periodos_m)} periodo(s).")

    def generate_url_candidates_mensual(year, month):
        candidates = []
        variantes_mes = [MESES[month]] + MESES_ALT.get(month, [])
        for mes_nombre in variantes_mes:
            carpeta_mes = f"{month:02d}_{mes_nombre}"
            for case_variant in {mes_nombre, mes_nombre.title(), mes_nombre.capitalize()}:
                archivo_zip = f"PMENSUAL_{case_variant}_{year}.zip"
                url = (
                    f"{BASE_URL}Operaci%C3%B3n%2FPrograma%20de%20Mantenimiento%2FPrograma%20Mensual%2F{year}%2F"
                    f"{carpeta_mes}%2FFinal%2F{archivo_zip}"
                )
                candidates.append(url)
        seen, unique = set(), []
        for c in candidates:
            if c not in seen:
                seen.add(c)
                unique.append(c)
        return unique

    def find_working_url(candidates, headers, timeout=20):
        tried = []
        for url in candidates:
            status = None
            try:
                resp = requests.head(url, headers=headers, timeout=timeout, allow_redirects=True)
                status = resp.status_code
                if status in (405, 501):
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

    def download_and_extract_mensual(year, month):
        mes_display = MESES[month]
        headers = {"User-Agent": "Mozilla/5.0"}
        candidates = generate_url_candidates_mensual(year, month)
        url, tried = find_working_url(candidates, headers)
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
                start_idx = col_letter_to_index(COL_START_MENSUAL)
                end_idx = col_letter_to_index(COL_END_MENSUAL)
                df = pd.read_excel(
                    io.BytesIO(excel_file.read()),
                    sheet_name=SHEET_MENSUAL,
                    header=HEADER_MENSUAL - 1,
                    usecols=list(range(start_idx, end_idx + 1)),
                )
                df = df.dropna(how="all").reset_index(drop=True)
                df["_year"] = year
                df["_month"] = month
                df["_periodo"] = f"{mes_display} {year}"
        return df, url, tried

    if st.button("🚀 Descargar y consolidar (Mensual)", type="primary", key="btn_mensual"):
        if not periodos_m:
            st.error("Selecciona al menos un año y un mes.")
        else:
            results, errors, debug_info = [], [], []
            progress = st.progress(0.0)
            status = st.empty()
            t0 = time.time()

            with ThreadPoolExecutor(max_workers=max_workers_m) as executor:
                futures = {
                    executor.submit(download_and_extract_mensual, y, m): (y, m)
                    for y, m in periodos_m
                }
                done = 0
                for future in as_completed(futures):
                    y, m = futures[future]
                    label = f"{MESES[m]} {y}"
                    try:
                        df, used_url, tried = future.result()
                        results.append(df)
                        debug_info.append({"periodo": label, "used_url": used_url, "tried": tried})
                    except Exception as e:
                        errors.append(f"{label}: {e}")
                    done += 1
                    progress.progress(done / len(futures))
                    status.text(f"Procesados {done}/{len(futures)} — último: {label}")

            elapsed = time.time() - t0
            status.empty()
            progress.empty()
            st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(periodos_m)} periodo(s), {max_workers_m} en paralelo)")

            if errors:
                with st.expander("⚠️ Errores", expanded=True):
                    for e in errors:
                        st.text(e)

            if show_debug_m and debug_info:
                with st.expander("🔎 Variantes de URL probadas", expanded=False):
                    for info in debug_info:
                        st.markdown(f"**{info['periodo']}** — usada: `{info['used_url']}`")
                        for u, s in info["tried"]:
                            st.text(f"  [{s}] {u}")

            if results:
                df_final_m = pd.concat(results, ignore_index=True)
                st.success(f"✅ {len(df_final_m):,} filas consolidadas de {len(results)} periodo(s).")
                st.dataframe(df_final_m.head(300), use_container_width=True)

                buf_xlsx = io.BytesIO()
                df_final_m.to_excel(buf_xlsx, index=False, engine="openpyxl")
                st.download_button(
                    "⬇️ Descargar PMENSUAL.xlsx", data=buf_xlsx.getvalue(),
                    file_name="PMENSUAL.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_xlsx_mensual",
                )
                buf_csv = df_final_m.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "⬇️ Descargar PMENSUAL.csv", data=buf_csv,
                    file_name="PMENSUAL.csv", mime="text/csv", key="dl_csv_mensual",
                )
            else:
                st.error("❌ No se consolidó ningún periodo.")


# ======================================================================
# TAB 2: PROGRAMA ANUAL (PAI) — LISTADO_Mantto
# ======================================================================
with tab_pai:
    st.markdown(
        """
    Descarga el archivo **LISTADO_Mantto** del Programa Anual de Mantenimiento (PAI),
    para uno o varios años, y consolida todo en una sola tabla.
    """
    )
    st.caption(
        "El nombre del ciclo y el prefijo del ZIP cambian cada año/ciclo del PAI "
        "(ej. '02_PAI Jul26-Jun27', 'SEGUNDO PAI') — ajústalos aquí si COES los cambia."
    )

    SHEET_PAI = "MANTTOS"
    HEADER_PAI = 9
    COL_RANGE_PAI = "B:P"

    col1, col2 = st.columns(2)
    with col1:
        ciclo_folder = st.text_input("Carpeta del ciclo PAI", value="02_PAI Jul26-Jun27", key="ciclo_pai")
        final_folder = st.text_input("Subcarpeta final", value="03_Final", key="final_pai")
    with col2:
        zip_prefix = st.text_input("Prefijo del ZIP", value="SEGUNDO PAI", key="prefijo_pai")
        max_workers_p = st.slider("Descargas en paralelo", 1, 6, 3, key="workers_pai")

    years_p = st.multiselect("Año(s)", options=list(range(2022, 2031)), default=[2026], key="years_pai")
    st.caption(f"Se descargará(n) {len(years_p)} año(s).")

    def generate_url_pai(year, ciclo_folder, final_folder, zip_prefix):
        archivo_zip = f"{zip_prefix}_{year}.zip"
        url = (
            f"{BASE_URL}Operaci%C3%B3n%2FPrograma%20de%20Mantenimiento%2FPrograma%20Anual%2F{year}%2F"
            f"{quote(ciclo_folder)}%2F{quote(final_folder)}%2F{quote(archivo_zip)}"
        )
        return url

    def download_and_extract_pai(year, ciclo_folder, final_folder, zip_prefix):
        url = generate_url_pai(year, ciclo_folder, final_folder, zip_prefix)
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=60)
        if response.status_code != 200:
            raise RuntimeError(f"HTTP {response.status_code} al descargar PAI {year} (url: {url})")

        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            namelist = z.namelist()
            excel_files = [
                f for f in namelist
                if "LISTADO_Mantto" in f and (f.endswith(".xlsx") or f.endswith(".xlsm"))
            ]
            if not excel_files:
                raise FileNotFoundError(
                    f"No se encontró LISTADO_Mantto en el ZIP del PAI {year} (url: {url}). "
                    f"Archivos disponibles: {namelist}"
                )
            excel_name = excel_files[0]
            with z.open(excel_name) as excel_file:
                df = pd.read_excel(
                    io.BytesIO(excel_file.read()),
                    sheet_name=SHEET_PAI,
                    header=HEADER_PAI - 1,
                    usecols=COL_RANGE_PAI,
                )
                df = df.dropna(how="all").reset_index(drop=True)
                df["AÑO"] = year
        return df, url

    if st.button("🚀 Descargar y consolidar (PAI)", type="primary", key="btn_pai"):
        if not years_p:
            st.error("Selecciona al menos un año.")
        else:
            results, errors = [], []
            progress = st.progress(0.0)
            status = st.empty()
            t0 = time.time()

            with ThreadPoolExecutor(max_workers=max_workers_p) as executor:
                futures = {
                    executor.submit(download_and_extract_pai, y, ciclo_folder, final_folder, zip_prefix): y
                    for y in years_p
                }
                done = 0
                for future in as_completed(futures):
                    y = futures[future]
                    try:
                        df, used_url = future.result()
                        results.append(df)
                    except Exception as e:
                        errors.append(f"Año {y}: {e}")
                    done += 1
                    progress.progress(done / len(futures))
                    status.text(f"Procesados {done}/{len(futures)} — último: {y}")

            elapsed = time.time() - t0
            status.empty()
            progress.empty()
            st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(years_p)} año(s), {max_workers_p} en paralelo)")

            if errors:
                with st.expander("⚠️ Errores", expanded=True):
                    for e in errors:
                        st.text(e)

            if results:
                df_final_p = pd.concat(results, ignore_index=True)
                st.success(f"✅ {len(df_final_p):,} filas consolidadas de {len(results)} año(s).")
                st.dataframe(df_final_p.head(300), use_container_width=True)

                buf_xlsx = io.BytesIO()
                df_final_p.to_excel(buf_xlsx, index=False, engine="openpyxl")
                st.download_button(
                    "⬇️ Descargar PAI.xlsx", data=buf_xlsx.getvalue(),
                    file_name="PAI.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_xlsx_pai",
                )
                buf_csv = df_final_p.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "⬇️ Descargar PAI.csv", data=buf_csv,
                    file_name="PAI.csv", mime="text/csv", key="dl_csv_pai",
                )
            else:
                st.error("❌ No se consolidó ningún año.")
