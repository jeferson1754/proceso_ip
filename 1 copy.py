from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill
from matplotlib.widgets import CheckButtons
import matplotlib.pyplot as plt
from matplotlib import patches
from datetime import datetime
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




def mostrar_grafico_historico(file_path):

    def leer_datos_desde_excel(archivo):
        xls = pd.ExcelFile(archivo)
        sheets = xls.sheet_names

        data_frames = []
        for sheet in sheets:
            # Extraer la fecha y hora del nombre de la hoja usando regex
            match = re.search(
                r'_(\d{4}-\d{2}-\d{2})_(\d{2}-\d{2}-\d{2})$', sheet)
            if match:
                fecha = match.group(1)
                hora = match.group(2).replace('-', ':')
                df = pd.read_excel(archivo, sheet_name=sheet)
                df['Fecha y Hora'] = pd.to_datetime(
                    f"{fecha} {hora}", format="%Y-%m-%d %H:%M:%S")
                data_frames.append(df)

        # Concatenar todos los DataFrames en uno solo
        all_data = pd.concat(data_frames, ignore_index=True)

        return all_data

    # Leer los datos desde el archivo Excel
    # Cambia esto por la ruta real de tu archivo Excel
    archivo = file_path
    df = leer_datos_desde_excel(archivo)

    # Convertir 'Conteo' a enteros
    df['Conteo'] = pd.to_numeric(
        df['Conteo'], errors='coerce').fillna(0).astype(int)

    # Verificar los datos después de la conversión
    # print("Datos después de la conversión a enteros:")
    # print(df.head())

    # Verificar si hay valores NaN en columnas importantes y limpiarlos
    df = df.dropna(subset=['Estado', 'Conteo'])

    # Crear el gráfico
    fig, ax = plt.subplots(figsize=(12, 6))

    # Ajusta los datos para el gráfico
    lines = []
    for estado in df['Estado'].unique():
        df_estado = df[df['Estado'] == estado]
        line, = ax.plot(df_estado['Fecha y Hora'],
                        df_estado['Conteo'], label=estado, marker='o')
        lines.append(line)

    plt.xlabel('Fecha y Hora')
    plt.ylabel('Conteo')
    plt.title('Equipos en la red')
    plt.legend()
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Crear la anotación que aparecerá al pasar el mouse
    annot = ax.annotate("", xy=(0, 0), xytext=(20, 20),
                        textcoords="offset points",
                        bbox=dict(boxstyle="round", fc="w"),
                        arrowprops=dict(arrowstyle="->"))
    annot.set_visible(False)

    # Función para actualizar la anotación
    def update_annot(line, ind):
        x, y = line.get_data()
        x_val = x[ind["ind"][0]]
        y_val = y[ind["ind"][0]]

        # Encuentra la fila en el DataFrame correspondiente a los datos de la línea
        df_estado = df[df['Estado'] == line.get_label()]
        fecha_hora = df_estado[df_estado['Fecha y Hora']
                               == x_val]['Fecha y Hora'].values[0]

        # Convertir numpy.datetime64 a datetime.datetime si es necesario
        if isinstance(fecha_hora, pd.Timestamp):
            fecha_hora = fecha_hora.to_pydatetime()
        elif isinstance(fecha_hora, np.datetime64):
            fecha_hora = pd.to_datetime(fecha_hora).to_pydatetime()

        # Formatear la fecha y hora en un formato más legible
        fecha_hora_formateada = fecha_hora.strftime("%Y-%m-%d %H:%M:%S")

        annot.xy = (x_val, y_val)
        text = f"{line.get_label()}\n{fecha_hora_formateada}\n{y_val}"
        annot.set_text(text)
        annot.get_bbox_patch().set_alpha(0.8)

    # Función para manejar los eventos del mouse

    def hover(event):
        vis = annot.get_visible()
        if event.inaxes == ax:
            for line in lines:
                cont, ind = line.contains(event)
                if cont:
                    update_annot(line, ind)
                    annot.set_visible(True)
                    fig.canvas.draw_idle()
                    return
        if vis:
            annot.set_visible(False)
            fig.canvas.draw_idle()

    fig.canvas.mpl_connect("motion_notify_event", hover)

    # Configurar los botones de selección
    # Posición en la parte superior del gráfico
    rax = plt.axes([0.4, 0.80, 0.2, 0.1])
    labels = df['Estado'].unique()
    visibility = [line.get_visible() for line in lines]
    check = CheckButtons(rax, labels, visibility)

    def func(label):
        for line in lines:
            if line.get_label() == label:
                line.set_visible(not line.get_visible())
        plt.draw()

    check.on_clicked(func)

    # Mostrar el gráfico
    print(f"Mostrando Graficos Historicos.")
    plt.show()



def main():


    excel_path = r"G:\Mi unidad\device_status_report.xlsx"


    mostrar_grafico_historico(excel_path)



if __name__ == "__main__":
    main()
