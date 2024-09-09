# Descripción: Este script automatiza el proceso de descarga de antecedentes disciplinarios de SIGA para los postulantes del día anterior.
import os, sys
from datetime import datetime, timedelta
import time
import pandas as pd
import datetime as dt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

from dotenv import load_dotenv

# Cargar las variables de entorno
load_dotenv()

# Variables de Entorno
urlSiga = os.getenv('urlsiga')
usernameSiga = os.getenv('usernamesiga')
passwordSiga = os.getenv('passwordidsiga')
urlSeccionSiga = os.getenv('urlSeccionSiga')
workfolder_path = os.getenv('workfolder_path')
RutaOnedrive = os.getenv('RutaOnedrive')
insumos_path = os.getenv('insumos_path')
chrome_driver_path = os.getenv('chrome_driver_path')

# Configuracion y Credenciales de envio Correos
username = os.getenv('SMTP_USERNAME')
password = os.getenv('SMTP_PASSWORD')
port = os.getenv('SMTP_PORT')
server = os.getenv('SMTP_SERVER')

# FUNCIONES PARA VALIDACION DEL SISTEMA
def creacion_carpeta():
    # Crear el directorio principal "AntecedentesSIGA" si no existe
    principal_folder = os.path.join(workfolder_path, "AntecedentesSIGA")
    if not os.path.exists(principal_folder):
        os.makedirs(principal_folder)

    fecha_actual = datetime.now().strftime('%Y-%m-%d')

    # Nombre de la carpeta con la fecha actual y el prefijo "AntecedentesDisciplinarios-"
    carpeta_nombre = "AntecedentesDisciplinarios-" + fecha_actual

    # Crear una carpeta con el nombre de la fecha actual dentro de "AntecedentesSIGA"
    fecha_folder = os.path.join(principal_folder, carpeta_nombre)
    if not os.path.exists(fecha_folder):
        os.makedirs(fecha_folder)
        print("Se creo la carpeta de la fecha actual")

    return fecha_folder

def Abrir_Navegador():
    # Opciones del navegador
    chrome_options = webdriver.ChromeOptions()

    # Llamar a la función 
    fecha_folder = creacion_carpeta()
    
    # Configurar las preferencias de un navegador web
    prefs = {"download.default_directory": fecha_folder, "plugins.always_open_pdf_externally": True}
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(executable_path=chrome_driver_path)

    #Abrir el navegador
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.refresh()

    #Limpiar las cookies
    driver.delete_all_cookies()

    #Ingresar a la pagina de SIGA
    driver.get(url=urlSiga)
    time.sleep(5)

    print("Abrir el navegador")
    
    return driver, fecha_folder

def enviar_correo(email_settings, msg):
    # Declaración de una variable para la fecha de hoy
    today = dt.datetime.today().strftime("%d-%m-%Y")

    # Configuración del formato del email
    message = MIMEMultipart()
    message['From'] = username
    message['To'] = ','.join(email_settings.get('recipients'))
    subject = email_settings.get('subject').replace('$(fecha)', today)
    message["Subject"] = subject
    message.attach(MIMEText(msg, 'plain'))
    
    # Envío de correo mediante una conexión SMTP
    try:
        with smtplib.SMTP(host=server, port=port, timeout=60) as conn:
            conn.starttls()
            conn.login(user=username, password=password)
            conn.sendmail(from_addr=username, to_addrs=email_settings.get('recipients'), msg=message.as_string())
    except Exception as e:
        print(f"Error al enviar el correo electronico: {e}")
    else:
        print("El Mensaje fue enviado con Exito")

def login_Verificacion():
    try:
        # Busca el Formulario de que este Activo y Presente
        FormLogin = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "/html/body/table/tbody/tr[3]/td/form")))
        print("Formulario Encontrado")
        login_Exitoso()
    except TimeoutException:
        # Si no se encuentra el formulario, enviar un correo electrónico
        email_settings = {'subject': 'Error en la Generación de Antecedentes SIGA $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com']}
        msg = f"""Error durante la ejecución,
        
        Motivo del error: El formulario en la página SIGA no se ha encontrado.
        
        Se requiere una revisión urgente de este error y validar su pronta resolución.
        
        Validar si su XPath {FormLogin} es correcto y si la página SIGA está disponible.

        Por favor, no responda ni envíe correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com

        Cordialmente, Automatizacion"""
        
        enviar_correo(email_settings, msg)
        return False
    else:
        return True

def login_Exitoso():
    try:
        # Validar si los campos del formulario esten e ingresar los datos de usuario
        login_usuario = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "/html/body/table/tbody/tr[3]/td/form/table/tbody/tr[1]/td/div/input")))
        login_usuario.send_keys(usernameSiga)
        login_contrasena = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "password")))
        login_contrasena.send_keys(passwordSiga)
        login_boton = driver.find_element(By.XPATH, "/html/body/table/tbody/tr[3]/td/form/table/tbody/tr[3]/td/table/tbody/tr/td/a")
        login_boton.click()
        print("Ingreso Exitosamente")
        
        Ingreso_A_Antecedentes()
        
        time.sleep(3)
        
    except Exception as e:
        print("Ocurrió un error: ", e)
        # Enviar un correo electrónico si el inicio de sesión no tiene éxito
        email_settings = {'subject': 'Error en la Generación de Antecedentes SIGA $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com']}

        msg = f"""Error durante la ejecución,
        
        Motivo del error: El intento de inicio de sesión en la página SIGA no ha tenido éxito.
        
        Se requiere una revisión urgente de este error y su pronta resolución.
        
        Validar si su XPath {login_usuario} y {login_contrasena} son correctos y si las credenciales de inicio de sesión son válidas.

        Por favor, no responda ni envíe correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com

        Cordialmente, Automatizacion"""
        
        enviar_correo(email_settings, msg)
        
        sys.exit(0)

def Ingreso_A_Antecedentes():
    try:
        url = urlSeccionSiga
        driver.get(url)
        print("Seccion de Antecedentes disciplinarios Encontrada Correctamente")
    except Exception as e:
        print("Ocurrió un error: ", e)
        # Enviar un correo electrónico si no se puede acceder a la sección de antecedentes disciplinarios
        email_settings = {'subject': 'Error en la Generación de Antecedentes SIGA $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com']}

        msg = """Error durante la ejecución,
        
        Motivo del error: No se pudo acceder correctamente a la sección 'AntecedentesDisciplinarios' en la página SIGA.
        
        Se requiere una revisión urgente de este error y su pronta resolución.

        Por favor, no responda ni envíe correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com

        Cordialmente, Automatizacion"""
        
        enviar_correo(email_settings, msg)
        sys.exit(0)

def obtener_Cedula_antecedentes():
    try:
        Resultados_path = []

        # Leer el archivo Excel y la hoja de trabajo
        df_documentos = pd.read_excel(insumos_path, sheet_name='Postulantes')

        # Convertir la columna 'Fecha de postulacion' a formato datetime
        df_documentos['Fecha de postulacion'] = pd.to_datetime(df_documentos['Fecha de postulacion'])
        #print(df_documentos['Fecha de postulacion'])

        # Obtener la fecha del día anterior
        #fecha_actual = pd.to_datetime("08/08/2024", format='%d/%m/%Y')
        #print(fecha_actual)
        
        fecha_actual = datetime.now()
        fecha_anterior = fecha_actual - timedelta(days=1)
        print(fecha_anterior.date())

        # Filtrar las filas donde 'Fecha de postulacion' es igual a la fecha actual
        df_documentos_hoy = df_documentos[df_documentos['Fecha de postulacion'].dt.date == fecha_anterior.date()]

        # Reemplazar los valores no finitos con 0 en la columna 'Documento'
        df_documentos_hoy['Documento '] = df_documentos_hoy['Documento '].fillna(0)

        # Convertir la columna 'Documento' a enteros
        df_documentos_hoy['Documento '] = df_documentos_hoy['Documento '].astype(int)

        # Eliminar las filas donde 'Documento' es 0
        df_documentos_hoy = df_documentos_hoy[df_documentos_hoy['Documento '] != 0]

        # Iterar sobre cada documento y añadirlo a la lista Resultados_path
        for documento in df_documentos_hoy['Documento ']:
            Resultados_path.append(documento)
        
        # Si no hay cedulas, se manda un Correo
        if not Resultados_path:
            print("No se encontraron cédulas para el día anterior")
            email_settings = {'subject': 'No se encontraron nuevos postulantes $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com', 'luisa.baez@gruporeditos.com']}

            msg = """ Hola!,

se le informa que, tras revisar el archivo Excel "PlanTalento.xlsx", no se han encontrado nuevos postulantes, 
Para Realizar el Proceso de Antecedentes Disciplinarios.

Por favor, no responda ni envíe correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com

Cordialmente, Automatizacion"""
            
            enviar_correo(email_settings, msg)
            
            time.sleep(5)
            
            # Cierra la Ejecucion
            sys.exit(0)

        documentos_sin_repetir = list(set(Resultados_path))
        return documentos_sin_repetir
    
    except Exception as e:
        print("Ocurrió un error: ", e)
        # Enviar un correo electrónico si no se pueden obtener las cédulas
        email_settings = {'subject': 'Error en la Generación de Antecedentes SIGA $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com']}

        msg = """Error durante la ejecución,
        
        Motivo del error: Intento de Obtener las Cedulas del Excel de PlanTalento.xlsx
        
        Se requiere una revisión urgente de este error y su pronta resolución.

        Por favor, no responda ni envíe correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com

        Cordialmente, Automatizacion"""
        
        enviar_correo(email_settings, msg)
        
        sys.exit(0)

def renombrar_archivo_descargado(cedula, download_path):
    # Esperar a que el archivo PDF aparezca en la carpeta de descargas
    timeout = 30
    seconds = 0
    archivo_pdf = None

    # Esperar a que el archivo PDF aparezca en la carpeta de descargas
    while seconds < timeout:
        for archivo in os.listdir(download_path):
            if archivo.endswith(".pdf") and "AntecedentesDisciplinariosPDF" in archivo:
                archivo_pdf = archivo
                break
        if archivo_pdf:
            break
        time.sleep(1)
        seconds += 1
    
    # Renombrar el archivo PDF con la cédula
    if archivo_pdf:
        nuevo_nombre = f"{cedula}.pdf"
        ruta_actual = os.path.join(download_path, archivo_pdf)
        nueva_ruta = os.path.join(download_path, nuevo_nombre)
        os.rename(ruta_actual, nueva_ruta)
        print(f"Archivo renombrado a {nuevo_nombre}")
    else:
        print(f"No se encontró el archivo PDF para la cédula {cedula} en el tiempo esperado")

def descargar_informe():
    # Llamar la Funcion Obtener Cedulas
    ruta_archivo_problemas = os.path.join(fecha_folder, 'cedulas_con_problemas.txt')

    try:
        ObtenerCedula = obtener_Cedula_antecedentes()
        
        # Iterar sobre cada cedula y descargar el informe
        for cedula in ObtenerCedula:
            try:
                Ingreso_A_Antecedentes()
                print(cedula)
            
                CampoAntecedentes = driver.find_element(By.ID, "ideUsuarioText")
                CampoAntecedentes.clear()
                CampoAntecedentes.send_keys(cedula)
                CampoAntecedentes.send_keys(Keys.ENTER)
            
                time.sleep(3)
            
                BtnGenerarReporte = driver.find_element(By.ID, "btnConsultar")
                BtnGenerarReporte.click()
            
                # Renombrar el archivo descargado con la cédula
                renombrar_archivo_descargado(cedula, fecha_folder)
                time.sleep(5)
            
                Ingreso_A_Antecedentes()
            except Exception as e:
                # Si hay una excepción, guarda la cédula en el archivo de texto
                with open(ruta_archivo_problemas, 'a') as file:
                    file.write(f'{cedula}\n')
                continue
            
    except Exception as e:
        print("Ocurrió un error: ", e)
        # Enviar un correo electrónico si no se pueden descargar los archivos PDF
        email_settings = {'subject': 'Error en la Generación de Antecedentes SIGA $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com']}

        msg = f"""Error durante la ejecución,
        
        Motivo del error: No se puedo Generar los Archivos PDF, Validar su Funcionamiento y dar su pronta Solución.
        
        Se requiere una revisión urgente de este error y su pronta resolución.
        
        Validar si su XPath {CampoAntecedentes} y {BtnGenerarReporte} son correctos y si las cédulas son válidas.

        Por favor, no responda ni envíe correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com

        Cordialmente, Automatizacion"""
        
        enviar_correo(email_settings, msg)
        sys.exit(0)

#Funcion que la carpeta creada, me haga una copia y la traslade a otra ruta
def trasladar_carpeta():
    # Directorio principal
    principal_folder = os.path.join(workfolder_path, "AntecedentesSIGA")
    Ruta_onedrive = os.path.join(RutaOnedrive, "AntecedentesSIGA")
    
    # Obtener la fecha del mes anterior
    fecha_actual = datetime.now().strftime('%Y-%m-%d')
    
    # Construir el prefijo de la carpeta del mes anterior
    prefijo_carpeta_mes_anterior = "AntecedentesDisciplinarios-"+fecha_actual
    
    # Listar todas las carpetas en el directorio principal
    carpetas = os.listdir(principal_folder)
    
    # Iterar sobre las carpetas y trasladar las que coincidan con el prefijo del mes anterior
    for carpeta in carpetas:
        if carpeta.startswith(prefijo_carpeta_mes_anterior):
            ruta_carpeta = os.path.join(principal_folder, carpeta)
            try:
                os.rename(ruta_carpeta, os.path.join(Ruta_onedrive, carpeta))
                print(f"Se trasladó la carpeta {ruta_carpeta}.")
            except OSError as e:
                print(f"No se pudo trasladar la carpeta {ruta_carpeta}: {e}")
        else:
            print(f"No se encontraron carpetas con el prefijo {prefijo_carpeta_mes_anterior}.")

# Inicio del Proceso
if __name__ == "__main__":
    print(datetime.now())
    print("Procesando")
    
    driver, fecha_folder = Abrir_Navegador()
    login_Verificacion()
    descargar_informe()
    
    time.sleep(5)
    trasladar_carpeta()
        
    print("Proceso Terminado Exitosamente")

    fecha_actual = datetime.now()
    fechafinal = fecha_actual.date()
    email_settings = {'subject': 'Se ejecuto Correctamente AntecedentesSIGA $(fecha)', 'recipients': ['aprendiz.funcional@gruporeditos.com', 'luisa.baez@gruporeditos.com']}
    msg = f"""Hola!, se le informa que la Automatizacion

Antecedentes SIGA, se ejecuto correctamente, por favor revisar la carpeta de la fecha {fechafinal}

por favor, no responda ni envíe correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com

Cordialmente, Automatizacion

"""
    enviar_correo(email_settings, msg)

    driver.quit()
    time.sleep(5)