from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from matplotlib.widgets import CheckButtons
import matplotlib.pyplot as plt
from matplotlib import patches
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import mplcursors
import subprocess
import pyautogui
import openpyxl
import logging
import time
import os
import re


def cargar_datos(file_path):
    """Cargar datos desde un archivo CSV, manejando diferentes codificaciones."""
    encodings = ['utf-16', 'ISO-8859-1']
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding,
                             delimiter='\t', on_bad_lines='skip')
            print(f"Datos cargados exitosamente del CSV.")
            return df
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f"Error al cargar el archivo '{file_path}': {e}")
        return None


def calcular_conteos(df):
    """Generar gráficos de barras, circular y de barras apiladas para el estado de los dispositivos."""
    if df is None or 'Estado' not in df.columns:
        print("Columna 'Estado' no encontrada en el DataFrame.")
        return None
    # Calcular las variables
    conteo_estado = df['Estado'].value_counts()
    conteo_estado_porcentaje = df['Estado'].value_counts(normalize=True)

    df['Segmento'] = df['IP'].apply(lambda x: '.'.join(
        x.split('.')[:3]) + '.0/24' if pd.notna(x) else 'Desconocido')
    conteo_segmentos_estados = df.groupby(
        ['Segmento', 'Estado']).size().unstack(fill_value=0)

    # Mostrar las variables en la consola
    '''
    print("Conteo de Estados:")
    print(conteo_estado)
    print("\nPorcentaje de Estados:")
    print(conteo_estado_porcentaje)
    print("\nConteo de Estados por Segmento:")
    print(conteo_segmentos_estados)
    '''
    return conteo_estado, conteo_estado_porcentaje, conteo_segmentos_estados


def formatear_hoja(ws, num_rows):
    """Formatear una hoja de cálculo de Excel."""
    # Alinear los porcentajes a la derecha
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=4, max_row=num_rows + 1):
        for cell in row:
            cell.alignment = Alignment(horizontal='right')

    # Alinear Totales y Porcentajes a la derecha
    for cell in ws['B'][num_rows - 2:num_rows]:  # Modificado para mayor claridad
        cell.alignment = Alignment(horizontal='right')

    # Ajustar el ancho de las columnas
    for column in ws.columns:
        max_length = max(len(str(cell.value)) for cell in column)
        ws.column_dimensions[column[0].column_letter].width = max_length + 2


def exportar_a_excel(conteo_segmentos_estados, excel_path):
    """Exportar datos a un archivo Excel, colocando la fecha actual primero."""
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # Crear o cargar el archivo Excel
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    # Crear o cargar la hoja de "Conteos"
    ws_conteos = wb.create_sheet(title=f"Conteos_{timestamp}", index=0)

    # Crear el DataFrame combinado
    df_conteos = pd.DataFrame({
        'Segmento': conteo_segmentos_estados.index,
        'Activado': conteo_segmentos_estados['Activado'],
        'Inactivo': conteo_segmentos_estados['Inactivo'],
        'Desconocido': conteo_segmentos_estados['Desconocido']
    })

    # Añadir la fila de totales
    df_conteos.loc['Totales'] = df_conteos[[
        'Activado', 'Inactivo', 'Desconocido']].sum()

    # Calcular los porcentajes
    total_activado = df_conteos['Activado'].sum()
    total_inactivo = df_conteos['Inactivo'].sum()
    total_desconocido = df_conteos['Desconocido'].sum()

    total_general = total_activado + total_inactivo + total_desconocido

    # Asegurarse de que los porcentajes sean calculados solo si hay un total mayor que cero
    if total_general > 0:
        porcentajes = [
            f"{total_activado / total_general * 100:.2f}%",
            f"{total_inactivo / total_general * 100:.2f}%",
            f"{total_desconocido / total_general * 100:.2f}%"
        ]
    else:
        porcentajes = ["0%", "0%", "0%"]

    # Añadir la fila de porcentajes, asegurando que los valores están en la columna correcta
    df_conteos.loc['Porcentajes'] = [
        None, porcentajes[0], porcentajes[1], porcentajes[2]]

    # Asegurar que las filas 'Totales' y 'Porcentajes' tengan su texto correspondiente
    df_conteos.at['Totales', 'Segmento'] = 'Totales'
    df_conteos.at['Porcentajes', 'Segmento'] = 'Porcentajes'

    # Añadir los datos a la hoja de "Conteos"
    for r in dataframe_to_rows(df_conteos, index=False, header=True):
        ws_conteos.append(r)

    # Formatear la hoja de "Conteos"
    formatear_hoja(ws_conteos, len(df_conteos))

    # Guardar el archivo Excel
    wb.save(excel_path)
    print(f"Datos exportados exitosamente a {excel_path}")


def main():

    file_path = r'C:\Users\jvargas\Documents\ip - copia.csv'
    excel_path = r"G:\Mi unidad\device_status_report - copia.xlsx"

    def procesar_datos(file_path, excel_path):
        # Cargar datos desde el archivo CSV
        df = cargar_datos(file_path)
        if df is not None:
            # Mostrar gráficos actuales
            conteo_estado, conteo_estado_porcentaje, conteo_segmentos_estados = calcular_conteos(
                df)

            # Exportar a Excel
            exportar_a_excel(conteo_segmentos_estados, excel_path)
        else:
            print("No se pudieron cargar los datos.")

        # Llamar a la función de procesamiento
    procesar_datos(file_path, excel_path)

    print(f"Proceso completado exitosamente")

    # Calcular la próxima hora
    proxima_hora = datetime.now() + timedelta(hours=2)
    print()  # Espacio en blanco
    print(f"Siguiente Escaneo a las {proxima_hora.strftime('%H:%M')} horas.")


if __name__ == "__main__":
    main()
