# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# Script para Subir Comentarios de Reunión desde Google Sheets a DBO
# Versión 1.4: 
# - Acción: Procesa Columna O (Índice 14), sube a DBO y limpia la celda.
# -----------------------------------------------------------------------------

import time
import os
import gspread
import google.auth
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURACIÓN DE CREDENCIALES ---
USUARIO_DBO = os.getenv('DBO_USUARIO', "mauro@dhartec.com.ar")
CONTRASENA_DBO = os.getenv('DBO_CONTRASENA', "94dBvyB5sv32^!$y")

# --- CONFIGURACIÓN DE URLs Y HOJAS ---
URL_LOGIN = "http://dbo2dhartec-env.eba-as23ttdp.us-west-2.elasticbeanstalk.com/spa/home"
URL_PERFORMANCE_BASE = "http://dbo2dhartec-env.eba-as23ttdp.us-west-2.elasticbeanstalk.com/spa/performances/{}/detail"

ID_REPORTE_DESTINO = "1hwlXkUE8fHQAl5OA29WGyg6xSnKKtETY1HfNNdUIcjM"
NOMBRE_HOJA_DESTINO = "Hoja 1"

# Índices de columnas (0-indexed)
COL_ID_IDX = 0        # Columna A
COL_COMENTARIO_IDX = 14 # Columna O

def inicializar_driver():
    """Configura Chrome en modo headless para GitHub Actions."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    service = ChromeService(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def subir_comentario_dbo(driver, wait, id_servicio, texto_original):
    """Navega a DBO y sube el comentario al servicio específico."""
    comentario_final = f"*Comentario de Reunion*: {texto_original}"
    print(f"  -> Procesando ID {id_servicio}...")
    try:
        driver.get(URL_PERFORMANCE_BASE.format(id_servicio))
        
        # Click en botón 'Crear' comentario
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Crear']"))).click()
        
        # Escribir en el campo de texto
        campo_texto = wait.until(EC.visibility_of_element_located((By.ID, "inputContent")))
        campo_texto.send_keys(comentario_final)
        
        # Click en 'Guardar'
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Guardar']"))).click()
        
        # Espera de 10s por lentitud del servidor DBO
        print(f"  -> Guardado. Esperando sincronización del servidor...")
        time.sleep(10)
        return True
    except Exception as e:
        print(f"    Error al subir comentario para ID {id_servicio}: {e}")
        return False

def main():
    print("Iniciando proceso de subida de comentarios de reunión...")
    
    # Autenticación con Google Sheets
    try:
        creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        GC = gspread.authorize(creds)
    except Exception as e:
        print(f"Error de autenticación con Google Sheets: {e}")
        return

    driver = inicializar_driver()
    wait = WebDriverWait(driver, 60)

    try:
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
        
        procesados = 0
        print("\n--- Escaneando Columna O para procesar ---")

        for i, row in enumerate(datos[1:], start=2): # Saltamos cabecera
            if len(row) > COL_COMENTARIO_IDX:
                id_servicio = row[COL_ID_IDX]
                texto_comentario = row[COL_COMENTARIO_IDX].strip()

                # Solo procesamos si hay ID y Comentario
                if id_servicio and texto_comentario:
                    if subir_comentario_dbo(driver, wait, id_servicio, texto_comentario):
                        # Limpiar celda O (gspread usa 1-based, así que es IDX + 1)
                        sheet.update_cell(i, COL_COMENTARIO_IDX + 1, "")
                        print(f"  [OK] Fila {i}: Subido y celda limpiada.")
                        procesados += 1
                        time.sleep(1) # Respetar límites de la API de Google

        if procesados == 0:
            print("No se encontraron comentarios pendientes en la columna O.")
        else:
            print(f"\nPrueba finalizada. Se procesaron {procesados} comentarios.")

    except Exception as e:
        print(f"Error crítico: {e}")
    finally:
        print("Cerrando navegador...")
        driver.quit()

if __name__ == "__main__":
    main()
