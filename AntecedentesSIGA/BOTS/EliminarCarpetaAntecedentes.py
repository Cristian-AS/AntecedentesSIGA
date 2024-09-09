# Descripción: Script para eliminar las carpetas de antecedentes disciplinarios del mes anterior.
import os
from datetime import datetime, timedelta

from dotenv import load_dotenv

# Cargar las variables de entorno
load_dotenv()

# Obtener la ruta de la carpeta de trabajo
RutaOnedrive  = os.getenv('RutaOnedrive')

# Función para eliminar las carpetas del mes anterior
def eliminar_carpetas_mes_anterior():
    # Directorio principal
    principal_folder = os.path.join(RutaOnedrive, "AntecedentesSIGA")

    # Obtener la fecha del mes anterior
    fecha_actual = datetime.now()
    primer_dia_mes_actual = fecha_actual.replace(day=1)
    fecha_mes_anterior = primer_dia_mes_actual - timedelta(days=1)
    nombre_mes_anterior = fecha_mes_anterior.strftime('%Y-%m')

    # Construir el prefijo de la carpeta del mes anterior
    prefijo_carpeta_mes_anterior = "AntecedentesDisciplinarios-" + nombre_mes_anterior

    # Listar todas las carpetas en el directorio principal
    carpetas = os.listdir(principal_folder)

    # Iterar sobre las carpetas y eliminar las que coincidan con el prefijo del mes anterior
    for carpeta in carpetas:
        if carpeta.startswith(prefijo_carpeta_mes_anterior):
            ruta_carpeta = os.path.join(principal_folder, carpeta)
            try:
                os.rmdir(ruta_carpeta)
                print(f"Se eliminó la carpeta {ruta_carpeta}.")
            except OSError as e:
                print(f"No se pudo eliminar la carpeta {ruta_carpeta}: {e}")

# Llamada a la función para eliminar las carpetas del mes anterior
eliminar_carpetas_mes_anterior()
