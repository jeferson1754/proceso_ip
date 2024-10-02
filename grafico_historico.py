import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.widgets import CheckButtons
from datetime import datetime
import numpy as np


def mostrar_grafico_historico(excel_path):

    def format_sheet_name(sheet_name):
        # Extraer la parte de la fecha del nombre de la hoja (después de "Conteos_")
        date_part = sheet_name.replace("Conteos_", "")
        # Convertir el string a un objeto datetime
        try:
            dt = datetime.strptime(date_part, "%Y-%m-%d_%H-%M-%S")
            return dt.strftime("%d-%m-%Y %H:%M:%S")
        except ValueError as e:
            print(f"Error parsing date for sheet '{sheet_name}': {e}")
            return sheet_name

    # Leer todas las hojas del archivo Excel
    xls = pd.ExcelFile(excel_path)
    all_totals = []

    for sheet_name in xls.sheet_names:
        # Leer cada hoja
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

        totals_row = df.iloc[-2].tolist()

        try:
            totals = {
                'Tiempo': format_sheet_name(sheet_name),
                'Activado': float(totals_row[1]) if isinstance(totals_row[1], (int, float)) else 0,
                'Inactivo': float(totals_row[2]) if isinstance(totals_row[2], (int, float)) else 0,
                'Desconocido': float(totals_row[3]) if isinstance(totals_row[3], (int, float)) else 0
            }
        except Exception as e:
            print(f"Error processing totals for sheet '{sheet_name}': {e}")
            continue

        all_totals.append(totals)

    df_totals = pd.DataFrame(all_totals)
    df_totals['Tiempo'] = pd.to_datetime(
        df_totals['Tiempo'], format='%d-%m-%Y %H:%M:%S')
    df_totals = df_totals.sort_values('Tiempo')

    # Crear el gráfico con layout restringido
    fig, ax = plt.subplots(figsize=(12, 6), constrained_layout=True)

    # Crear las líneas con marcadores
    lines = []
    lines.append(ax.plot(df_totals['Tiempo'], df_totals['Activado'],
                         marker='o', label='Activado', linewidth=2, markersize=8)[0])
    lines.append(ax.plot(df_totals['Tiempo'], df_totals['Inactivo'],
                         marker='o', label='Inactivo', linewidth=2, markersize=8)[0])
    lines.append(ax.plot(df_totals['Tiempo'], df_totals['Desconocido'],
                         marker='o', label='Desconocido', linewidth=2, markersize=8)[0])

    # Personalizar el gráfico
    ax.set_title('Evolución histórica de totales de dispositivos',
                 fontsize=14, pad=20)
    ax.set_xlabel('Tiempo', fontsize=12)
    ax.set_ylabel('Número de dispositivos', fontsize=12)
    ax.grid(True, linestyle='--', alpha=0.7)
    ax.legend(fontsize=10)

    # Rotar y ajustar las etiquetas del eje x para mejor legibilidad
    plt.xticks(rotation=45, ha='right')

    # Configurar las anotaciones
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

        # Convertir el valor y en un entero
        y_val_int = int(y_val)

        # Encuentra la fila en el DataFrame correspondiente a los datos de la línea
        df_estado = df_totals[df_totals['Tiempo'] == x_val]

        # Formatear la fecha y hora
        fecha_hora_formateada = pd.to_datetime(
            x_val).strftime("%Y-%m-%d %H:%M:%S")

        # Actualizar la anotación
        annot.xy = (x_val, y_val)
        text = f"{line.get_label()}\n{fecha_hora_formateada}\n{y_val_int}"
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
    rax = plt.axes([0.4, 0.80, 0.2, 0.1])  # Posición del panel de checkboxes
    labels = [line.get_label() for line in lines]
    visibility = [line.get_visible() for line in lines]
    check = CheckButtons(rax, labels, visibility)

    # Función para alternar la visibilidad de las líneas

    def func(label):
        for line in lines:
            if line.get_label() == label:
                line.set_visible(not line.get_visible())
        plt.draw()

    check.on_clicked(func)

    # Rotar y ajustar las etiquetas del eje x para mejor legibilidad
    plt.xticks(rotation=45, ha='right')

    # Ajustar los márgenes para evitar que se corten las etiquetas
    plt.tight_layout()
    
    # Mostrar el gráfico
    print("Mostrando Gráfico Histórico.")
    plt.show()


def main():

    excel_path = r"G:\Mi unidad\device_status_report 2.xlsx"

    mostrar_grafico_historico(excel_path)

    print(f"Proceso completado exitosamente")
