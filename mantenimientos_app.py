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


def fetch_valid_zip(candidates, headers, timeout=30):
    """
    Prueba cada URL candidata con un GET en streaming (COES no soporta bien HEAD
    en este endpoint de descarga, así que no lo usamos) y devuelve la primera que
    sea un ZIP real: status 200, content-type que no sea HTML, y bytes iniciales 'PK'.
    Devuelve (url, contenido_bytes, lista_de_intentos) o (None, None, lista_de_intentos).
    """
    tried = []
    for url in candidates:
        try:
            resp = requests.get(url, headers=headers, timeout=timeout, stream=True)
        except Exception as e:
            tried.append((url, f"ERROR: {e}"))
            continue

        status = resp.status_code
        content_type = resp.headers.get("Content-Type", "")

        if status != 200:
            tried.append((url, status))
            resp.close()
            continue

        if "html" in content_type.lower():
            tried.append((url, f"200 pero content-type={content_type} (no es un ZIP real)"))
            resp.close()
            continue

        content = resp.content  # descarga completa solo del candidato que llegó hasta aquí
        resp.close()

        if content[:2] != b"PK":
            snippet = content[:200].decode("latin-1", errors="ignore").replace("\n", " ")
            tried.append((url, f"200 pero no son bytes de ZIP (inicio: {snippet!r})"))
            continue

        tried.append((url, status))
        return url, content, tried

    return None, None, tried


tab_mensual, tab_preliminar, tab_pai = st.tabs([
    "📅 Programa Mensual (Intervenciones)",
    "📋 Programa Mensual Preliminar (Agentes)",
    "📆 Programa Anual (PAI)",
])

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

    def download_and_extract_mensual(year, month):
        mes_display = MESES[month]
        headers = {"User-Agent": "Mozilla/5.0"}
        candidates = generate_url_candidates_mensual(year, month)
        url, content, tried = fetch_valid_zip(candidates, headers)
        if url is None:
            detail = "\n".join(f"  [{status}] {u}" for u, status in tried)
            raise RuntimeError(
                f"No se encontró un ZIP válido para {mes_display} {year} tras probar "
                f"{len(tried)} variante(s):\n{detail}"
            )

        with zipfile.ZipFile(io.BytesIO(content)) as z:
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
# TAB 2: PROGRAMA MENSUAL PRELIMINAR (Anexo1_Intervenciones_(Agentes), versión Preliminar)
# ======================================================================
with tab_preliminar:
    st.markdown(
        """
    Descarga el archivo **Anexo1_Intervenciones_(Agentes)** del Programa Mensual, tanto en su
    versión **Final** como **Preliminar** (útil cuando el mes en curso aún no tiene versión Final).
    """
    )

    SHEET_PRE = "MANTTOS"
    HEADER_PRE = 9
    COL_START_PRE, COL_END_PRE = "B", "Q"

    with st.expander("⚙️ Opciones", expanded=False):
        max_workers_pre = st.slider("Descargas en paralelo", 1, 6, 3, key="workers_preliminar")
        show_debug_pre = st.checkbox("Mostrar variantes de URL probadas (debug)", key="debug_preliminar")

    col1, col2, col3 = st.columns(3)
    with col1:
        years_pre = st.multiselect("Año(s)", options=list(range(2022, 2031)), default=[2026], key="years_preliminar")
    with col2:
        months_pre = st.multiselect(
            "Mes(es)", options=list(MESES.keys()), default=[8],
            format_func=lambda m: MESES[m], key="months_preliminar",
        )
    with col3:
        stage_pre = st.selectbox("Versión", options=["Preliminar", "Final"], index=0, key="stage_preliminar")

    periodos_pre = [(y, m) for y in sorted(years_pre) for m in sorted(months_pre)]
    st.caption(f"Se descargarán {len(periodos_pre)} periodo(s) — versión {stage_pre}.")

    def generate_url_candidates_preliminar(year, month, stage):
        variantes_mes = [MESES[month]] + MESES_ALT.get(month, [])
        candidates = []
        for mes_nombre in variantes_mes:
            carpeta_mes = f"{month:02d}_{mes_nombre}"
            for case_variant in {mes_nombre, mes_nombre.title(), mes_nombre.capitalize()}:
                if stage.upper() == "FINAL":
                    carpeta_stage_opts = ["Final"]
                    archivo_zip_opts = [f"PMENSUAL_{case_variant}_{year}.zip"]
                else:
                    carpeta_stage_opts = ["Preliminar"]
                    archivo_zip_opts = [
                        f"PMI_{case_variant}_{year}_PRELIMINAR.zip",
                        f"PMI_{case_variant}_{year}_Preliminar.zip",
                    ]
                for carpeta_stage in carpeta_stage_opts:
                    for archivo_zip in archivo_zip_opts:
                        url = (
                            f"{BASE_URL}Operaci%C3%B3n%2FPrograma%20de%20Mantenimiento%2FPrograma%20Mensual%2F{year}%2F"
                            f"{carpeta_mes}%2F{carpeta_stage}%2F{archivo_zip}"
                        )
                        candidates.append(url)
        seen, unique = set(), []
        for c in candidates:
            if c not in seen:
                seen.add(c)
                unique.append(c)
        return unique

    def download_and_extract_preliminar(year, month, stage):
        mes_display = MESES[month]
        headers = {"User-Agent": "Mozilla/5.0"}
        candidates = generate_url_candidates_preliminar(year, month, stage)
        url, content, tried = fetch_valid_zip(candidates, headers)
        if url is None:
            detail = "\n".join(f"  [{status}] {u}" for u, status in tried)
            raise RuntimeError(
                f"No se encontró un ZIP válido para {mes_display} {year} ({stage}) tras probar "
                f"{len(tried)} variante(s):\n{detail}"
            )

        with zipfile.ZipFile(io.BytesIO(content)) as z:
            namelist = z.namelist()
            excel_files = [
                f for f in namelist
                if "INTERVENCIONES_(AGENTES)" in f.upper().replace(" ", "_")
                and f.endswith((".xlsx", ".xlsm", ".xls"))
            ]
            if not excel_files:
                raise FileNotFoundError(
                    f"No se encontró Anexo1_Intervenciones_(Agentes) en el ZIP de {mes_display} {year} "
                    f"({stage}) (url usada: {url}). Archivos disponibles: {namelist}"
                )
            excel_name = excel_files[0]
            with z.open(excel_name) as excel_file:
                start_idx = col_letter_to_index(COL_START_PRE)
                end_idx = col_letter_to_index(COL_END_PRE)
                df = pd.read_excel(
                    io.BytesIO(excel_file.read()),
                    sheet_name=SHEET_PRE,
                    header=HEADER_PRE - 1,
                    usecols=list(range(start_idx, end_idx + 1)),
                )
                df = df.dropna(how="all").reset_index(drop=True)
                df["PERIODO"] = f"{mes_display}_{year}"
                df["ESTADO_PROGRAMA"] = stage.upper()
        return df, url, tried

    if st.button("🚀 Descargar y consolidar (Preliminar)", type="primary", key="btn_preliminar"):
        if not periodos_pre:
            st.error("Selecciona al menos un año y un mes.")
        else:
            results, errors, debug_info = [], [], []
            progress = st.progress(0.0)
            status = st.empty()
            t0 = time.time()

            with ThreadPoolExecutor(max_workers=max_workers_pre) as executor:
                futures = {
                    executor.submit(download_and_extract_preliminar, y, m, stage_pre): (y, m)
                    for y, m in periodos_pre
                }
                done = 0
                for future in as_completed(futures):
                    y, m = futures[future]
                    label = f"{MESES[m]} {y} ({stage_pre})"
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
            st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(periodos_pre)} periodo(s), {max_workers_pre} en paralelo)")

            if errors:
                with st.expander("⚠️ Errores", expanded=True):
                    for e in errors:
                        st.text(e)

            if show_debug_pre and debug_info:
                with st.expander("🔎 Variantes de URL probadas", expanded=False):
                    for info in debug_info:
                        st.markdown(f"**{info['periodo']}** — usada: `{info['used_url']}`")
                        for u, s in info["tried"]:
                            st.text(f"  [{s}] {u}")

            if results:
                df_final_pre = pd.concat(results, ignore_index=True)
                st.success(f"✅ {len(df_final_pre):,} filas consolidadas de {len(results)} periodo(s).")
                st.dataframe(df_final_pre.head(300), use_container_width=True)

                buf_xlsx = io.BytesIO()
                df_final_pre.to_excel(buf_xlsx, index=False, engine="openpyxl")
                st.download_button(
                    "⬇️ Descargar PMENSUAL_preliminar.xlsx", data=buf_xlsx.getvalue(),
                    file_name="PMENSUAL_preliminar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_xlsx_preliminar",
                )
                buf_csv = df_final_pre.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "⬇️ Descargar PMENSUAL_preliminar.csv", data=buf_csv,
                    file_name="PMENSUAL_preliminar.csv", mime="text/csv", key="dl_csv_preliminar",
                )
            else:
                st.error("❌ No se consolidó ningún periodo.")


# ======================================================================
# TAB 3: PROGRAMA ANUAL (PAI) — LISTADO_Mantto
# ======================================================================
with tab_pai:
    st.markdown(
        """
    Descarga el archivo **LISTADO_Mantto** del Programa Anual de Mantenimiento (PAI),
    para uno o varios años, y consolida todo en una sola tabla.
    """
    )
    st.caption(
        "Cada año COES publica dos ciclos de PAI: uno **Enero-Diciembre** (mismo año) "
        "y otro **Julio-Junio** (cruza al año siguiente). El nombre de la carpeta se "
        "arma automáticamente a partir del año que elijas."
    )

    SHEET_PAI = "MANTTOS"
    HEADER_PAI = 9
    COL_RANGE_PAI = "B:Q"

    PAI_CYCLES = {
        1: {"label": "01 - PAI Enero-Diciembre (mismo año)", "prefix": "PRIMER PAI"},
        2: {"label": "02 - PAI Julio-Junio (cruza al año siguiente)", "prefix": "SEGUNDO PAI"},
    }

    def build_ciclo_folder_variants(year: int, cycle: int):
        """Devuelve variantes del nombre de carpeta del ciclo, porque COES a veces
        le agrega un espacio inicial raro al nombre (confirmado en el ciclo 1 de 2026)."""
        yy = str(year)[2:]
        if cycle == 1:
            base = f"01_PAI Ene{yy}-Dic{yy}"
            return [f" {base}", base]  # variante con espacio inicial primero (caso confirmado)
        else:
            yy2 = str(year + 1)[2:]
            base = f"02_PAI Jul{yy}-Jun{yy2}"
            return [base, f" {base}"]

    def generate_url_candidates_pai(year, cycle, final_override="", prefix_override=""):
        ciclo_folder_variants = build_ciclo_folder_variants(year, cycle)
        default_prefix = PAI_CYCLES[cycle]["prefix"]

        prefixes = [prefix_override] if prefix_override else [default_prefix, default_prefix.replace(" ", "_")]
        final_folders = [final_override] if final_override else ["03_Final", "02_Final", "Final", "01_Final"]
        separators = ["_", " "]  # antes del año: "PRIMER PAI_2026.zip" vs "PRIMER PAI 2026.zip"

        candidates = []
        for ciclo_folder in ciclo_folder_variants:
            for ff in final_folders:
                for pfx in prefixes:
                    for sep in separators:
                        archivo_zip = f"{pfx}{sep}{year}.zip"
                        url = (
                            f"{BASE_URL}Operaci%C3%B3n%2FPrograma%20de%20Mantenimiento%2FPrograma%20Anual%2F{year}%2F"
                            f"{quote(ciclo_folder)}%2F{quote(ff)}%2F{quote(archivo_zip)}"
                        )
                        candidates.append(url)

        seen, unique = set(), []
        for c in candidates:
            if c not in seen:
                seen.add(c)
                unique.append(c)
        return unique

    col1, col2 = st.columns(2)
    with col1:
        years_p = st.multiselect("Año(s)", options=list(range(2022, 2031)), default=[2026], key="years_pai")
    with col2:
        cycles_p = st.multiselect(
            "Ciclo(s) PAI",
            options=list(PAI_CYCLES.keys()),
            default=[1, 2],
            format_func=lambda c: PAI_CYCLES[c]["label"],
            key="cycles_pai",
        )

    with st.expander("⚙️ Opciones avanzadas", expanded=False):
        max_workers_p = st.slider("Descargas en paralelo", 1, 6, 3, key="workers_pai")
        show_debug_p = st.checkbox("Mostrar variantes de URL probadas (debug)", key="debug_pai")
        st.caption("Si la detección automática falla, puedes forzar la subcarpeta final y/o el prefijo del ZIP:")
        final_folder_override = st.text_input(
            "Forzar subcarpeta final (dejar vacío = probar automáticamente)", value="", key="final_override_pai"
        )
        prefix_override = st.text_input(
            "Forzar prefijo del ZIP (dejar vacío = usar 'PRIMER PAI' / 'SEGUNDO PAI')",
            value="", key="prefix_override_pai",
        )

    periodos_p = [(y, c) for y in sorted(years_p) for c in sorted(cycles_p)]
    st.caption(f"Se descargarán {len(periodos_p)} combinación(es) de año/ciclo.")

    def download_and_extract_pai(year, cycle, final_override, prefix_override):
        cycle_label = PAI_CYCLES[cycle]["label"]
        headers = {"User-Agent": "Mozilla/5.0"}
        candidates = generate_url_candidates_pai(year, cycle, final_override, prefix_override)
        url, content, tried = fetch_valid_zip(candidates, headers)
        if url is None:
            detail = "\n".join(f"  [{status}] {u}" for u, status in tried)
            raise RuntimeError(
                f"No se encontró un ZIP válido para {cycle_label} {year} tras probar "
                f"{len(tried)} variante(s):\n{detail}"
            )

        with zipfile.ZipFile(io.BytesIO(content)) as z:
            namelist = z.namelist()
            excel_files = [
                f for f in namelist
                if "LISTADO_Mantto" in f and (f.endswith(".xlsx") or f.endswith(".xlsm"))
            ]
            if not excel_files:
                raise FileNotFoundError(
                    f"No se encontró LISTADO_Mantto en el ZIP de {cycle_label} {year} (url: {url}). "
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
        return df, url, tried

    if st.button("🚀 Descargar y consolidar (PAI)", type="primary", key="btn_pai"):
        if not periodos_p:
            st.error("Selecciona al menos un año y un ciclo.")
        else:
            results, errors, debug_info = [], [], []
            progress = st.progress(0.0)
            status = st.empty()
            t0 = time.time()

            with ThreadPoolExecutor(max_workers=max_workers_p) as executor:
                futures = {
                    executor.submit(
                        download_and_extract_pai, y, c, final_folder_override, prefix_override
                    ): (y, c)
                    for y, c in periodos_p
                }
                done = 0
                for future in as_completed(futures):
                    y, c = futures[future]
                    label = f"{PAI_CYCLES[c]['label']} — {y}"
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
            st.caption(f"⏱️ Tiempo total: {elapsed:.1f} s ({len(periodos_p)} combinación(es), {max_workers_p} en paralelo)")

            if errors:
                with st.expander("⚠️ Errores", expanded=True):
                    for e in errors:
                        st.text(e)

            if show_debug_p and debug_info:
                with st.expander("🔎 Variantes de URL probadas", expanded=False):
                    for info in debug_info:
                        st.markdown(f"**{info['periodo']}** — usada: `{info['used_url']}`")
                        for u, s in info["tried"]:
                            st.text(f"  [{s}] {u}")

            if results:
                df_final_p = pd.concat(results, ignore_index=True)
                st.success(f"✅ {len(df_final_p):,} filas consolidadas de {len(results)} combinación(es).")
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
                st.error("❌ No se consolidó ninguna combinación de año/ciclo.")
