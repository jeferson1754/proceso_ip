import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

ESTADO_ORDER = ['Activado', 'Inactivo', 'Desconocido']
COLOR_MAP = {'Desconocido': 'red',
             'Inactivo': 'orange',
             'Activado': 'green'}


def cargar_datos_estados(excel_path, sheet_index=0):
    """
    Cargar datos de Excel usando índices de columnas.

    Args:
        excel_path (str): Ruta al archivo Excel
        sheet_index (int): Índice de la hoja a cargar

    Returns:
        pd.DataFrame: DataFrame con los datos cargados
    """
    try:
        # Cargar el DataFrame sin nombres de columnas
        df = pd.read_excel(excel_path, sheet_name=sheet_index, header=None)

        # Filtrar las filas que no contienen "Segmento", "Total" o "Porcentaje"
        df = df[~df[0].astype(str).str.contains(
            'Segmento|Total|Porcentaje', case=False, na=False)]

        # Establecer el segmento como índice
        df.set_index(0, inplace=True)

        # Convertir valores a numéricos
        for col in [1, 2, 3]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(
                    ',', ''), errors='coerce').fillna(0)

        return df
    except Exception as e:
        print(f"Error al cargar los datos: {e}")
        return pd.DataFrame()


def _calcular_cambio_porcentual(actual, anterior):
    """Calcula el cambio porcentual entre dos valores."""
    return float('inf') if anterior == 0 and actual > 0 else ((actual - anterior) / anterior) * 100


def _generar_texto_comparativo(estado, actual, anterior):
    """Genera texto comparativo entre valores actuales y anteriores."""
    cambio = _calcular_cambio_porcentual(actual, anterior)
    if abs(cambio) < 1:
        return f"{estado}: Sin cambios significativos"
    direccion = "más" if cambio > 0 else "menos"
    return f"{abs(cambio):.1f}% {direccion} que el anterior"


def mostrar_graficos(df_actual, df_anterior=None):
    """
    Generar gráficos para el estado de los dispositivos usando índices de columnas.

    Args:
        df_actual (pd.DataFrame): DataFrame actual con datos
        df_anterior (pd.DataFrame, optional): DataFrame con datos anteriores
    """

    # Crear la figura con más espacio vertical
    fig = plt.figure(figsize=(16, 14))

    # Definir la disposición de los subplots con más espacio
    gs = fig.add_gridspec(3, 2, height_ratios=[1, 0.1, 1.2], hspace=0.4)

    # Crear los subplots
    ax1 = fig.add_subplot(gs[0, 0])  # Gráfico de barras
    ax2 = fig.add_subplot(gs[0, 1])  # Gráfico circular
    # Gráfico de barras apiladas (ocupa todo el ancho)
    ax3 = fig.add_subplot(gs[2, :])

    # Preparar datos para los gráficos
    conteo_estado = pd.Series({
        'Activado': df_actual[1].sum(),
        'Inactivo': df_actual[2].sum(),
        'Desconocido': df_actual[3].sum()
    })

    if df_anterior is not None:
        conteo_anterior = pd.Series({
            'Activado': df_anterior[1].sum(),
            'Inactivo': df_anterior[2].sum(),
            'Desconocido': df_anterior[3].sum()
        })
    else:
        conteo_anterior = None

    # Asignar los colores a cada estado
    colors = [COLOR_MAP[estado] for estado in ESTADO_ORDER]

    # Gráfico de barras simple
    bars = ax1.bar(range(len(ESTADO_ORDER)), [
                   conteo_estado[estado] for estado in ESTADO_ORDER], color=colors)
    ax1.set_title('Conteo de Dispositivos por Estado', pad=20)
    ax1.set_xlabel('')
    ax1.set_ylabel('Número de Dispositivos')
    ax1.set_xticks(range(len(ESTADO_ORDER)))
    ax1.set_xticklabels(ESTADO_ORDER, rotation=0)

    # Añadir etiquetas a las barras
    max_height = max(conteo_estado)
    for i, v in enumerate(conteo_estado):
        ax1.text(i, v, str(int(v)), ha='center', va='bottom')

        # Agregar comparación si hay datos anteriores
        if conteo_anterior is not None:
            cambio = _generar_texto_comparativo(
                ESTADO_ORDER[i], v, conteo_anterior[ESTADO_ORDER[i]])
            y_pos = v - max_height * 0.1
            ax1.text(i, y_pos, cambio, ha='center', va='top', fontsize=8,
                     color='blue', bbox=dict(facecolor='white', edgecolor='blue', alpha=0.7, pad=3))

    # Ajustar límites del eje y para dar más espacio
    ax1.set_ylim(0, max_height * 1.2)

    # Gráfico circular
    ax2.axis('off')  # Utiliza el cuadrado existente

    ax2 = fig.add_axes([0.5, 0.5, 0.4, 0.4])  # Cambia la posición y el tamaño
    wedges, texts, autotexts = ax2.pie(
        [conteo_estado[estado] for estado in ESTADO_ORDER],
        labels=ESTADO_ORDER,
        colors=colors,
        autopct='%1.1f%%'
    )

    ax2.set_title('Distribución de Dispositivos por Estado')

    # Ajustar el tamaño del texto en el gráfico circular
    plt.setp(autotexts, size=8)
    plt.setp(texts, size=8)

    # Gráfico de barras apiladas
    data_stacked = {
        'Activado': df_actual[1],
        'Inactivo': df_actual[2],
        'Desconocido': df_actual[3]
    }
    df_stacked = pd.DataFrame(data_stacked)

    # Crear gráfico apilado
    bars = df_stacked.plot(kind='bar', stacked=True, ax=ax3, color=colors)

    ax3.set_title('Conteo de Dispositivos por Segmento de IP y Estado', pad=20)
    ax3.set_xlabel('Segmento de IP')
    ax3.set_ylabel('Número de Dispositivos')
    ax3.set_xticklabels(df_actual.index, rotation=45, ha='right')

    # Configurar los límites del eje x
    ax3.set_xlim(-0.5, len(df_actual.index) - 0.5)

    # Ajustar los límites del eje y
    ylim = ax3.get_ylim()
    ax3.set_ylim(ylim[0], ylim[1] * 1.1)

    # Configurar tooltip interactivo
    annot = ax3.annotate("", xy=(0, 0), xytext=(20, 20), textcoords="offset points",
                         bbox=dict(boxstyle="round", fc="lightyellow",
                                   ec="orange", alpha=0.8, pad=0.5),
                         arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=.2", color="orange"))
    annot.set_visible(False)

    def update_annot(bar, segmento, estado, valor):
        x = bar.get_x() + bar.get_width() / 2.
        y = bar.get_y() + bar.get_height() / 2.

        xytext = (20, -20) if y > ax3.get_ylim()[1] * 0.7 else (20, 20)

        annot.xyann = xytext
        annot.xy = (x, y)
        text = f"{estado}: {valor}\n{segmento}"
        annot.set_text(text)

        color_map = {'Activado': 'lightgreen',
                     'Inactivo': 'lightsalmon', 'Desconocido': 'lightgray'}
        annot.get_bbox_patch().set_facecolor(color_map[estado])
        annot.get_bbox_patch().set_alpha(0.9)

    def hover(event):
        vis = annot.get_visible()
        if event.inaxes == ax3:
            for i, estado_bars in enumerate(bars.containers):
                for j, bar in enumerate(estado_bars):
                    if bar.contains(event)[0]:
                        segmento = df_actual.index[j]
                        estado = ESTADO_ORDER[i]
                        valor = int(bar.get_height())
                        update_annot(bar, segmento, estado, valor)
                        annot.set_visible(True)
                        fig.canvas.draw_idle()
                        return
        if vis:
            annot.set_visible(False)
            fig.canvas.draw_idle()

    fig.canvas.mpl_connect("motion_notify_event", hover)

    # Configurar leyenda
    ax3.legend(ESTADO_ORDER, title='Estado',
               loc='center left', bbox_to_anchor=(1, 0.5))

    # Ajustar el diseño manualmente
    plt.subplots_adjust(top=0.85, bottom=0.1, left=0.1,
                        right=0.85, hspace=0.4, wspace=0.4)

    # Ajustar layout para evitar superposiciones
    # plt.tight_layout()
    print("Mostrando Gráficos de Datos Recientes.")

    plt.show()


def visualizar_estados(excel_path):
    """
    Función principal para ejecutar la visualización de estados.

    Args:
        excel_path (str): Ruta al archivo Excel con los datos
    """
    # Cargar datos actuales y anteriores
    df_actual = cargar_datos_estados(excel_path, sheet_index=0)
    df_anterior = cargar_datos_estados(excel_path, sheet_index=1)

    if df_actual.empty:
        print("No se pudieron cargar los datos actuales correctamente.")
        return

    # Mostrar gráficos
    mostrar_graficos(df_actual, df_anterior)
    print("Visualización completada exitosamente.")
