# -*- coding: utf-8 -*-
"""
Proyecto: Automatización para filtrar ventas en Excel
Autor: Atziri Alejandra Rodríguez Peña
Fecha: Julio 2025

Descripción:
Este script lee un archivo Excel con datos de ventas,
filtra las ventas que se pagaron en efectivo y exporta el resultado
a un archivo CSV para facilitar su análisis.
"""

import pandas as pd
import os

def leer_archivo():
    carpeta_entrada = "input"
    archivo = input("👉 Ingresa el nombre del archivo Excel (incluye la extensión, ej: datos.xlsx): ")
    ruta_completa = os.path.join(carpeta_entrada, archivo)

    print("🔄 Leyendo archivo Excel...")
    columnas_interes = [2, 3, 4, 5, 6, 12]  # Columnas relevantes para el análisis
    try:
        df = pd.read_excel(ruta_completa, sheet_name='Sheet1', header=0, usecols=columnas_interes)
        print(f"✅ Archivo '{archivo}' cargado correctamente.")
        return df
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{archivo}' en la carpeta '{carpeta_entrada}'. Revisa el nombre e intenta de nuevo.")
        exit(1)
    except Exception as e:
        print(f"❌ Ocurrió un error al leer el archivo: {e}")
        exit(1)

def filtrar_datos(df):
    print("🔍 Filtrando solo las ventas pagadas en efectivo ('Cash')...")
    df_filtrado = df[df['Payment'] == 'Cash']
    print(f"➡️ Se encontraron {len(df_filtrado)} registros con pago en efectivo.")
    return df_filtrado

def exportar_csv(df):
    carpeta_salida = "output"
    archivo_salida = "resultado.csv"
    ruta_salida = os.path.join(carpeta_salida, archivo_salida)
    print(f"💾 Exportando datos filtrados a '{ruta_salida}'...")
    try:
        df.to_csv(ruta_salida, sep=',', index=False, header=True)
        print("✅ Exportación completada con éxito.")
    except Exception as e:
        print(f"❌ Error al exportar el archivo CSV: {e}")
        exit(1)

def main():
    df = leer_archivo()
    df_filtrado = filtrar_datos(df)
    exportar_csv(df_filtrado)
    input("👍 Proceso finalizado. Presiona Enter para salir...")

if __name__ == '__main__':
    main()
