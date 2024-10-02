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



def main():

    excel_path = r"G:\Mi unidad\device_status_report 2.xlsx"

    #mostrar_grafico_historico2(excel_path)

    print(f"Proceso completado exitosamente")


if __name__ == "__main__":
    main()
