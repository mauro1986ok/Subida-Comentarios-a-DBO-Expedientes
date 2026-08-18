# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# Script para Subir Comentarios de Reunión desde Google Sheets a DBO
# Versión 1.7: 
# - Acción: Procesa Columna N (Índice 13), sube a DBO y limpia la celda.
# - Corrección: Cambiada ruta de 'performances' a 'documents' para expedientes.
# - Integración: Alertas a Google Chat y variables de entorno (.env).
# -----------------------------------------------------------------------------

import time
import os
import datetime
import gspread
import google.auth
import requests
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Cargar variables de entorno
load_dotenv()

# --- CONFIGURACIÓN DE CREDENCIALES ---
USUARIO_DBO = os.getenv('DBO_USUARIO')
CONTRASENA_DBO = os.getenv('DBO_CONTRASENA')
GCHAT_WEBHOOK = os.getenv('GCHAT_WEBHOOK')

# --- CONFIGURACIÓN DE URLs Y HOJAS ---
URL_LOGIN = "http://dbo2dhartec-env.eba-as23ttdp.us-west-2.elasticbeanstalk.com/spa/home"
URL_DOCUMENTO_BASE = "http://dbo2dhartec-env.eba-as23ttdp.us-west-2.elasticbeanstalk.com/spa/documents/{}/detail"

ID_REPORTE_DESTINO = "1TOHmtI8wFsegTPwmTFPP9UzLQS5IVvX6yPyhC9LOdQU"
NOMBRE_HOJA_DESTINO = "Expedientes"

# Índices de columnas (0-indexed)
COL_ID_IDX = 0          # Columna A (ID)
COL_COMENTARIO_IDX = 13 # Columna N (Comentario de reunión)

def notificar_chat(tipo, mensaje_detalle):
    """Envía un reporte a Google Chat con formato estructurado."""
    if not GCHAT_WEBHOOK:
        print("Advertencia: No se encontró el Webhook de Google Chat en el entorno.")
        return

    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    iconos = {
        "ERROR": "🚨 *ERROR GRAVE*",
        "FALLA_CARGA": "⚠️ *FALLA DE CARGA*",
        "INFO": "ℹ️ *ACTUALIZACIÓN/RESUMEN*"
    }
    
    encabezado = iconos.get(tipo, "🔔 *NOTIFICACIÓN*")
    
    mensaje_formateado = (
        f"{encabezado}\n"
        f"📂 *Script:* `meeting_uploader.py` (Expedientes)\n"
        f"📅 *Fecha:* {fecha_actual}\n"
        f"🌐 *Entorno:* GitHub Actions\n"
        f"📝 *Detalle:* {mensaje_detalle}"
    )
    
    try:
        requests.post(GCHAT_WEBHOOK, json={"text": mensaje_formateado})
    except Exception as e:
        print(f"No se pudo enviar la notificación a Google Chat: {e}")

def inicializar_driver():
    """Configura Chrome en modo headless para GitHub Actions."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    service = ChromeService(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def subir_comentario_dbo(driver, wait, id_expediente, texto_original):
    """Navega a DBO y sube el comentario al expediente específico."""
    comentario_final = f"*Comentario de Reunion*: {texto_original}"
    print(f"  -> Procesando Expediente ID {id_expediente}...")
    try:
        # Navegamos a la nueva ruta de documentos
        driver.get(URL_DOCUMENTO_BASE.format(id_expediente))
        
        # Click en botón 'Crear' comentario
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Crear']"))).click()
        
        # Escribir en el campo de texto
        campo_texto = wait.until(EC.visibility_of_element_located((By.ID, "inputContent")))
        campo_texto.send_keys(comentario_final)
        
        # Click en 'Guardar'
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar']"))).click()
        
        # Espera de 10s por lentitud del servidor DBO
        print(f"  -> Guardado exitoso. Esperando sincronización...")
        time.sleep(10)
        return True
    except Exception as e:
        error_msg = f"Error al subir comentario para Expediente {id_expediente}: {str(e)[:200]}"
        print(f"    {error_msg}")
        notificar_chat("FALLA_CARGA", error_msg)
        return False

def main():
    print("Iniciando proceso de subida de comentarios de reunión a Expedientes...")
    
    # Autenticación con Google Sheets
    try:
        creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        GC = gspread.authorize(creds)
    except Exception as e:
        error_msg = f"Error de autenticación con Google Sheets: {e}"
        print(error_msg)
        notificar_chat("ERROR", error_msg)
        return

    driver = None
    errores_carga = 0
    procesados = 0

    try:
        driver = inicializar_driver()
        wait = WebDriverWait(driver, 60)

        # 1. Login en DBO
        print("Iniciando sesión en DBO...")
        driver.get(URL_LOGIN)
        wait.until(EC.visibility_of_element_located((By.ID, "username"))).send_keys(USUARIO_DBO)
        driver.find_element(By.ID, "password").send_keys(CONTRASENA_DBO)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "nav.navbar")))
        print("Sesión DBO iniciada.")

        # 2. Leer Hoja de Cálculo
        sheet = GC.open_by_key(ID_REPORTE_DESTINO).worksheet(NOMBRE_HOJA_DESTINO)
        datos = sheet.get_all_values()
        
        print("\n--- Escaneando Columna N para procesar comentarios ---")

        for i, row in enumerate(datos[1:], start=2): # Saltamos cabecera
            if len(row) > COL_COMENTARIO_IDX:
                id_expediente = row[COL_ID_IDX]
                texto_comentario = row[COL_COMENTARIO_IDX].strip()

                # Solo procesamos si hay ID y Comentario en la columna objetivo
                if id_expediente and texto_comentario:
                    if subir_comentario_dbo(driver, wait, id_expediente, texto_comentario):
                        # Limpiar celda N (gspread usa 1-based)
                        sheet.update_cell(i, COL_COMENTARIO_IDX + 1, "")
                        print(f"  [OK] Fila {i}: Comentario subido al expediente y celda limpiada.")
                        procesados += 1
                        time.sleep(1) # Respetar límites de la API de Google
                    else:
                        errores_carga += 1

        # Resumen final de la ejecución
        if procesados > 0 or errores_carga > 0:
            resumen = f"Proceso finalizado. Éxitos: {procesados}. Fallos: {errores_carga}."
            print(resumen)
            notificar_chat("INFO", resumen)
        else:
            print("No se encontraron comentarios pendientes en la columna N.")

    except Exception as e:
        error_msg = f"Error crítico durante la ejecución web: {e}"
        print(error_msg)
        notificar_chat("ERROR", error_msg)
    finally:
        print("Cerrando navegador...")
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
