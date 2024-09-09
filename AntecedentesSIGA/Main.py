# Description: Script principal que ejecuta los bloques de código de los bots de SIGA y Antecedentes
import os
import sys
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()
workfolder_path = os.getenv('workfolder_path')

# Obtener la fecha actual
fecha_actual = datetime.now()

def CrearConsoleLog(Path):
    # Definir la ruta de la carpeta de log
    log_path = os.path.join(Path, 'Console_log')

    # Crear la carpeta de log si no existe
    if not os.path.exists(log_path):
        os.mkdir(log_path)
        print(f"Se ha creado la carpeta de log en {log_path}")

    return log_path

original_stdout = sys.stdout

# Formatear la fecha y la hora como una cadena
timestamp = fecha_actual.strftime("%Y-%m-%d_%H-%M-%S")

# Definir la ruta del archivo de console log
log_path = CrearConsoleLog(workfolder_path)

with open(os.path.join(log_path, f'console_log_{timestamp}.txt'), 'a') as f:
    sys.stdout = f # Cambiar la salida estándar al archivo que acabamos de abrir

    # Cambiar el directorio de trabajo al directorio del script
    os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'BOTS'))

    try:
        # Verificar si es el día 2 del mes
        if fecha_actual.day == 2:
            # Ejecutar el Segundo bloque
            print(f"Es el día 2 del mes. Fecha actual: {fecha_actual.date()}")
            exec(open('EliminarCarpetaAntecedentes.py', encoding='utf-8').read())
        else:
            print(f"No es el día 2 del mes. Fecha actual: {fecha_actual.date()}")

        # Ejecutar el Primer bloque
        exec(open('IngresarSIGA.py', encoding='utf-8').read())

    except Exception as e:
        print(f"Error: {e.message}")
    
    sys.stdout = original_stdout # Restaurar la salida estándar original



