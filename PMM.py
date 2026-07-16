# -*- coding: utf-8 -*-
import requests
import io
import zipfile
import pandas as pd

RUTA_ARCHIVO = r"C:\Users\GZ6710\OneDrive - ENGIE\Escritorio\ENGIE\2026\Plexos\SCRIPT\PMENSUAL_2026.xlsx"

# Nombres de mes en español tal como aparecen en las carpetas de COES
MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SETIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}


def generate_url(year, month):
    """
    Construye la URL de descarga del Programa Mensual para un año y mes dados.
    Ej: 2026, 7 ->
    .../Programa Mensual/2026/07_JULIO/Final/PMENSUAL_JULIO_2026.zip
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


def download_and_extract(year, month):
    url = generate_url(year, month)
    mes_nombre = MESES[month]
    print(f"Descargando ZIP para {mes_nombre} {year}...")
    print(f"URL: {url}")

    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"Error al descargar: status {response.status_code}")
        return None

    try:
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            namelist = z.namelist()

            # Archivo de interés: Anexo1_Intervenciones_(Agentes)_<MES>_<AÑO>.xlsm
            excel_files = [
                f for f in namelist
                if 'INTERVENCIONES_(AGENTES)' in f.upper().replace(' ', '_')
                and f.endswith(('.xlsx', '.xlsm', '.xls'))
            ]

            if not excel_files:
                print("No se encontró el archivo Anexo1_Intervenciones_(Agentes) dentro del ZIP.")
                print("Archivos disponibles:", namelist)
                return None

            excel_name = excel_files[0]
            print(f"Leyendo archivo: {excel_name}")

            with z.open(excel_name) as excel_file:
                df = pd.read_excel(
                    io.BytesIO(excel_file.read()),
                    sheet_name='MANTTOS',
                    header=8,                     # fila 9 -> header=8 (0-indexed)
                    usecols=list(range(1, 17))    # columnas B a Q (índices 1 a 16)
                )
                df = df.dropna(how='all').reset_index(drop=True)

                print("Preview:")
                print(df.head())
                return df

    except Exception as e:
        print(f"Error procesando el ZIP de {mes_nombre} {year}: {e}")
        return None


def main(periodos):
    """
    periodos: lista de tuplas (year, month), ej. [(2026, 7)]
    """
    all_data = []

    for year, month in periodos:
        df = download_and_extract(year, month)
        if df is not None and not df.empty:
            all_data.append(df)

    if all_data:
        consolidated_df = pd.concat(all_data, ignore_index=True)
        print("\nDataFrame consolidado:")
        print(consolidated_df.head())

        # Sobrescribe el archivo solo con los datos de esta corrida
        # (NO se concatena con lo que ya hubiera en RUTA_ARCHIVO)
        consolidated_df.to_excel(RUTA_ARCHIVO, index=False)
        print(f"\nDatos guardados exitosamente en '{RUTA_ARCHIVO}'")
    else:
        print("No se procesó ningún dato.")


if __name__ == "__main__":
    # Ejemplo: descargar solo Julio 2026
    periodos = [(2026, 7)]
    main(periodos)