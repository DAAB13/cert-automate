import pandas as pd
import config 
import funciones
import os
import locale

try:
  locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
  try:
    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')
  except locale.Error:
    print("ADVERTENCIA: No se pudo establecer el idioma a español.")

def procesar_solicitudes():
  try:    #--- FASE 1: LEER ARCHIVO EXCEL ---#
    df = pd.read_excel(
      config.RUTA_EXCEL,
      sheet_name="cert-cons"
    )
    procesar_solicitudes = df[
      (~df['Código'].isin(config.EXCLUIR_VALORES)) & (df['Estado'].isna())
    ].dropna(subset=['Año'])

    if procesar_solicitudes.empty:
      print("👍No se encontraron nuevas solicitudes pendientes")
    else:
      print(f"🔥Se encontraton {len(procesar_solicitudes)} solicitud(es) por enviar")
    # recorremos el dataframe con filtro, una fila a la vez
    for index, fila in procesar_solicitudes.iterrows():
      print(f"▶️ Procesando solicitudes de: {fila['Nombres']} (Fila original: {index + 2})")
      ruta_pdf = funciones.generar_documento(fila) # capturamos la ruta del pdf
      #verificamos si el doc se creó corrextamente antes de continuar
      if ruta_pdf and os.path.exists(ruta_pdf):
        funciones.crear_qr_firmar(fila, ruta_pdf)
      else:
        print(f"❌ No se pudo procesar la solicitud para {fila['Nombres']} debido a un error en la generación del documento.")
  except FileNotFoundError: 
    print(f"❌ ¡ERROR CRÍTICO! No se pudo encontrar el archivo Excel en la ruta:")
    print(f"   {config.RUTA_EXCEL}")
  except Exception as e:
    print(f"❌ ¡ERROR INESPERADO! El programa se detuvo por: {e}")

if __name__ == "__main__":
    procesar_solicitudes()