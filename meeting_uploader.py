# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# Script para Subir Comentarios de Reunión desde Google Sheets a DBO
# Versión 2.0:
# - Migración: Reemplazado Selenium por la API REST oficial de DBO (DBOClient).
# - Acción: Procesa Columna N (Índice 13), sube a DBO vía API y limpia la celda.
# - Integración: Alertas a Google Chat y variables de entorno (.env).
# -----------------------------------------------------------------------------

import os
import time
import datetime
import gspread
import google.auth
import requests
from dotenv import load_dotenv

from dbo_client import DBOClient

# Cargar variables de entorno
load_dotenv()

# --- CONFIGURACIÓN DE NOTIFICACIONES ---
GCHAT_WEBHOOK = os.getenv('GCHAT_WEBHOOK')

# --- CONFIGURACIÓN DE HOJAS DE GOOGLE SHEETS ---
ID_REPORTE_DESTINO = os.getenv('ID_REPORTE_DESTINO', "1TOHmtI8wFsegTPwmTFPP9UzLQS5IVvX6yPyhC9LOdQU")
NOMBRE_HOJA_DESTINO = os.getenv('NOMBRE_HOJA_DESTINO', "Expedientes")

# Índices de columnas (0-indexed)
COL_ID_IDX = 0          # Columna A (ID)
COL_COMENTARIO_IDX = 13 # Columna N (Comentario de reunión)


def notificar_chat(tipo: str, mensaje_detalle: str):
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
        f"📂 *Script:* `meeting_uploader.py` (Expedientes API)\n"
        f"📅 *Fecha:* {fecha_actual}\n"
        f"🌐 *Entorno:* GitHub Actions / Local\n"
        f"📝 *Detalle:* {mensaje_detalle}"
    )
    
    try:
        requests.post(GCHAT_WEBHOOK, json={"text": mensaje_formateado}, timeout=10)
    except Exception as e:
        print(f"No se pudo enviar la notificación a Google Chat: {e}")


def subir_comentario_dbo(client: DBOClient, id_expediente: str, texto_original: str) -> bool:
    """Navega a DBO vía API y sube el comentario al expediente específico."""
    print(f"  -> Procesando Expediente ID {id_expediente}...")
    try:
        client.expediente_crear_comentario(
            document_id=id_expediente,
            text=texto_original,
            prefix="*Comentario de Reunion*:"
        )
        print(f"  -> Guardado exitoso en DBO vía API.")
        return True
    except Exception as e:
        error_msg = f"Error al subir comentario para Expediente {id_expediente}: {str(e)[:200]}"
        print(f"    {error_msg}")
        notificar_chat("FALLA_CARGA", error_msg)
        return False


def main():
    print("Iniciando proceso de subida de comentarios de reunión a Expedientes (API DBO)...")
    
    # 1. Autenticación con Google Sheets
    try:
        creds, _ = google.auth.default(scopes=[
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ])
        GC = gspread.authorize(creds)
    except Exception as e:
        error_msg = f"Error de autenticación con Google Sheets: {e}"
        print(error_msg)
        notificar_chat("ERROR", error_msg)
        return

    # 2. Inicialización y Login con DBOClient
    try:
        client = DBOClient.from_env()
        print("Sesión DBO iniciada exitosamente vía API.")
    except Exception as e:
        error_msg = f"Error al inicializar/autenticar con la API de DBO: {e}"
        print(error_msg)
        notificar_chat("ERROR", error_msg)
        return

    errores_carga = 0
    procesados = 0

    try:
        # 3. Leer Hoja de Cálculo
        sheet = GC.open_by_key(ID_REPORTE_DESTINO).worksheet(NOMBRE_HOJA_DESTINO)
        datos = sheet.get_all_values()
        
        print("\n--- Escaneando Columna N para procesar comentarios ---")

        for i, row in enumerate(datos[1:], start=2):  # Saltamos cabecera
            if len(row) > COL_COMENTARIO_IDX:
                id_expediente = str(row[COL_ID_IDX]).strip() if row[COL_ID_IDX] else ""
                texto_comentario = str(row[COL_COMENTARIO_IDX]).strip()

                # Solo procesamos si hay ID y Comentario en la columna objetivo
                if id_expediente and texto_comentario:
                    if subir_comentario_dbo(client, id_expediente, texto_comentario):
                        # Limpiar celda N (gspread usa índices 1-based)
                        sheet.update_cell(i, COL_COMENTARIO_IDX + 1, "")
                        print(f"  [OK] Fila {i}: Comentario subido al expediente {id_expediente} y celda limpiada.")
                        procesados += 1
                        time.sleep(1)  # Respetar límites de la API de Google Sheets
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
        error_msg = f"Error crítico durante la ejecución: {e}"
        print(error_msg)
        notificar_chat("ERROR", error_msg)
    finally:
        print("Cerrando sesión de API DBO...")
        client.close()


if __name__ == "__main__":
    main()
