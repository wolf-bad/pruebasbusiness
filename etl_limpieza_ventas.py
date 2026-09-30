# -*- coding: utf-8 -*-
"""
================================================================================
 ETL DE LIMPIEZA - VENTAS DE SUCURSALES (TIENDA DE TECNOLOGIA)
================================================================================
 Archivo fuente  : ventas_sucursales.xlsx   (NO se modifica, solo se lee)
 Archivo destino : ventas_limpias.xlsx      (datos limpios)
 Registro nulos  : registros_nulos.txt      (filas eliminadas por valores nulos)

 Transformaciones realizadas (pipeline ETL):
   1. EXTRACT  : Lectura del Excel original sin alterarlo.
   2. TRIM     : Eliminacion de espacios en blanco al inicio/final de cada campo
                 de texto (y normalizacion de espacios internos multiples).
   3. NORMALIZ.: Fecha convertida al formato dd/mm/yyyy; sucursal con Formato
                 Título; ids y cantidad como enteros.
   4. DUPLICADOS: Eliminacion de filas completamente duplicadas (keep='first').
   5. NULOS    : Filas con valores nulos se separan y se guardan en un .txt;
                 luego se eliminan del dataset limpio.
   6. CARGA    : Guardado del resultado en ventas_limpias.xlsx.
================================================================================
"""

import os
import pandas as pd

# ------------------------------------------------------------------------------
# CONFIGURACION DE RUTAS (relativas a la ubicacion de este script)
# ------------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_ORIGINAL = os.path.join(BASE_DIR, "ventas_sucursales.xlsx")   # solo lectura
ARCHIVO_LIMPIO   = os.path.join(BASE_DIR, "ventas_limpias.xlsx")      # salida Excel
ARCHIVO_NULOS    = os.path.join(BASE_DIR, "registros_nulos.txt")      # salida TXT


# ==============================================================================
# 1) EXTRACT - Extraer (leer) los datos crudos del Excel ORIGINAL
# ==============================================================================
def extract():
    """Lee el archivo ventas_sucursales.xlsx tal cual, sin modificarlo.
    Todo se lee como texto (dtype=str) para poder aplicar trim antes de
    convertir tipos, evitando que pandas interprete mal las fechas."""
    df = pd.read_excel(ARCHIVO_ORIGINAL, dtype=str)
    print(f"[EXTRACT]  {len(df)} filas leidas desde '{os.path.basename(ARCHIVO_ORIGINAL)}'")
    return df


# ==============================================================================
# 2) TRANSFORM - TRIM: limpiar espacios de TODOS los campos
# ==============================================================================
def limpiar_espacios(df):
    """Aplica strip() a cada celda de texto: elimina espacios al inicio y al
    final (' Santa Cruz ' -> 'Santa Cruz') y colapsa espacios internos
    multiples en uno solo. Las celdas vacias quedan como NaN."""
    for columna in df.columns:
        df[columna] = (
            df[columna]
            .astype("string")            # tipo de texto que soporta NaN
            .str.strip()                 # TRIM: quita espacios izq./der.
            .str.replace(r"\s+", " ", regex=True)  # un solo espacio interno
            .replace("", pd.NA)          # cadenas vacias -> valor nulo
        )
    print("[TRANSFORM] TRIM aplicado a todos los campos")
    return df


# ==============================================================================
# 2b) TRANSFORM - NORMALIZACION DE TIPOS Y FORMATO DE FECHA dd/mm/yyyy
# ==============================================================================
def normalizar_tipos_y_fecha(df):
    """Convierte cada campo a su tipo definitivo:
       - fecha       -> datetime (formato dd/mm/yyyy) y vuelta a string
                        'dd/mm/yyyy'. Valores no parseables -> NaT (nulo).
       - id_venta, id_cliente, id_producto, cantidad -> enteros (Int64).
       - sucursal    -> Formato Titulo ('LA PAZ' -> 'La Paz'), estandarizando
                        el nombre de las sucursales de la tienda."""
    # --- FECHA: forzar lectura dd/mm/yyyy (dayfirst=True) ------------------
    df["fecha"] = pd.to_datetime(df["fecha"], format="%d/%m/%Y", dayfirst=True,
                                 errors="coerce")
    # --- SUCURSAL: nombres propios en Formato Titulo ------------------------
    df["sucursal"] = df["sucursal"].str.title()
    # --- NUMERICOS: ids y cantidades como enteros ----------------------------
    for col in ["id_venta", "id_cliente", "id_producto", "cantidad"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    # --- Devolver la fecha como texto dd/mm/yyyy en el Excel de salida -------
    df["fecha"] = df["fecha"].dt.strftime("%d/%m/%Y")
    print("[TRANSFORM] Tipos normalizados y fecha en formato dd/mm/yyyy")
    return df


# ==============================================================================
# 2c) TRANSFORM - ELIMINAR DUPLICADOS
# ==============================================================================
def eliminar_duplicados(df):
    """Elimina filas totalmente duplicadas conservando la primera aparicion
    (keep='first'). Se hace ANTES de quitar nulos para no duplicar trabajo."""
    antes = len(df)
    df = df.drop_duplicates(keep="first").reset_index(drop=True)
    print(f"[TRANSFORM] Duplicados eliminados: {antes - len(df)} fila(s)")
    return df


# ==============================================================================
# 2d) TRANSFORM - SEPARAR Y ELIMINAR FILAS CON VALORES NULOS
# ==============================================================================
def separar_nulos(df):
    """Identifica las filas que contienen ALGUN valor nulo (NaN/NA), las
    guarda en un archivo de texto (registros_nulos.txt) como bitacora de
    auditoria y las elimina del dataset limpio."""
    mask_nulos = df.isna().any(axis=1)          # True si la fila tiene nulos
    df_nulos = df[mask_nulos].copy()            # filas rechazadas (auditoria)
    df_limpio = df[~mask_nulos].reset_index(drop=True)  # filas validas

    # --- Guardar las filas nulas en un archivo TXT -------------------------
    with open(ARCHIVO_NULOS, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write("BITACORA DE REGISTROS NULOS - ETL ventas_sucursales.xlsx\n")
        f.write("Filas eliminadas por contener uno o mas valores nulos\n")
        f.write("=" * 70 + "\n\n")
        f.write(f"Total de filas con valores nulos: {len(df_nulos)}\n\n")
        if not df_nulos.empty:
            # to_string() conserva la estructura de tabla para revision manual
            f.write(df_nulos.to_string(index=False) + "\n")
        else:
            f.write("No se encontraron filas con valores nulos.\n")

    print(f"[TRANSFORM] Filas con nulos eliminadas: {len(df_nulos)} "
          f"(guardadas en '{os.path.basename(ARCHIVO_NULOS)}')")
    return df_limpio


# ==============================================================================
# 3) LOAD - Cargar (guardar) los datos limpios en un NUEVO Excel
# ==============================================================================
def load(df):
    """Escribe ventas_limpias.xlsx con una hoja 'Ventas_Limpias', ajustando
    el ancho de columnas automaticamente. El original NO se toca."""
    with pd.ExcelWriter(ARCHIVO_LIMPIO, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ventas_Limpias")
        hoja = writer.sheets["Ventas_Limpias"]
        for i, col in enumerate(df.columns, start=1):
            ancho = max(len(str(col)), df[col].astype(str).str.len().max()) + 3
            hoja.column_dimensions[hoja.cell(row=1, column=i).column_letter].width = ancho
    print(f"[LOAD]      Datos limpios guardados en '{os.path.basename(ARCHIVO_LIMPIO)}'")


# ==============================================================================
# PIPELINE ETL - orquesta todo el flujo Extract -> Transform -> Load
# ==============================================================================
def pipeline_etl():
    df = extract()                     # 1. EXTRAER del Excel original
    filas_iniciales = len(df)          # guarda el conteo para el resumen final
    df = limpiar_espacios(df)          # 2. TRIM de todos los campos
    df = normalizar_tipos_y_fecha(df)  # 3. Tipos + fecha dd/mm/yyyy
    df = eliminar_duplicados(df)       # 4. Eliminar duplicados
    df = separar_nulos(df)             # 5. Nulos -> TXT y fuera del dataset
    load(df)                           # 6. CARGAR nuevo Excel limpio
    print("\n--- RESUMEN FINAL ---")
    print(f"Filas iniciales : {filas_iniciales}")
    print(f"Filas limpias   : {len(df)}")
    print(f"Columnas        : {list(df.columns)}")
    return df


if __name__ == "__main__":
    resultado = pipeline_etl()
    print("\nDatos finales limpios:")
    print(resultado.to_string(index=False))
