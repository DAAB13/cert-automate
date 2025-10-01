import config
import re
import os
import qrcode
import fitz
from datetime import datetime
from docx.shared import Pt
from docxtpl import DocxTemplate
from docxtpl import RichText
from docx2pdf import convert


def obtener_superindice(dia):
  if 11 <= dia <= 13:
    return "th"
  elif dia % 10 == 1:
    return "st"
  elif dia % 10 == 2:
    return "nd"
  elif dia % 10 == 3:
    return "rd"
  else:
    return "th"
  

def formatear_fecha(fecha_x, idioma, detalle_servicio):
  if "Examen de comprensión de textos" in detalle_servicio:
    mes_es = fecha_x.strftime("%B")
    return {'texto_simple': f"{fecha_x.day} de {mes_es} del {fecha_x.year}"}
  
  elif idioma == "Portugués":
    mes_es = fecha_x.strftime("%B")
    mes_portugues = config.MES_PORTUGUES.get(mes_es.lower(), mes_es)
    return {'texto_simple': f"em {fecha_x.day} de {mes_portugues} de {fecha_x.year}"}

  elif idioma == "Inglés":
    superindice = obtener_superindice(fecha_x.day)
    mes_es = fecha_x.strftime("%B")
    mes_ingles = config.MES_INGLES.get(mes_es.lower(), mes_es)
    return {
      'dia': fecha_x.day,
      'superindice_dia': superindice,
      'mes': mes_ingles,
      'año': fecha_x.year
    }
  else:
    # Formato para español (por defecto)
    mes_es = fecha_x.strftime("%B")
    return {'texto_simple': f"el {fecha_x.day} de {mes_es} de {fecha_x.year}"}
  


def procesar_nivel(fila, idioma, detalle_servicio):
  if "Certificado de nivel" in detalle_servicio: 
    nivel_original = fila["Indica el nivel culminado"]
    if idioma != "Español":
      nivel_traducido = config.NIVEL_TRADUCCION.get(idioma.lower(), {}).get(nivel_original.lower(), nivel_original)
      return {'descripcion': nivel_traducido.capitalize(), 'mce': None}
    return {'descripcion': nivel_original.capitalize(), 'mce': None}
  if "Examen de Suficiencia" in detalle_servicio:
    resultado_examen = fila["Resultado examen o curso"]
    #re.search = Regex (expresiones regulares)
    #la usamos para encontrar la primera coincidencia del patrón en el texto
    match = re.search(r"([a-zA-ZÁ-ÿ\s]+?)\s+([A-Z0-9]+)$", resultado_examen) #re.search(patron, texto)
    if match: #si match es verdadero 
      descripcion, mce = match.groups() #.groups devuelve una tupla con strings diferentes
      if idioma != "Español":
        resultado_traducido = config.NIVEL_TRADUCCION.get(idioma.lower(), {}).get(descripcion.strip().lower(), descripcion.strip())
        return {'descripcion': resultado_traducido.capitalize(), 'mce': mce}
      return {'descripcion': descripcion.stri().capitalize(), 'mce': mce}
    return {'descripcion': resultado_examen, 'mce': None} 
  if "Examen de comprensión de textos" in detalle_servicio:
    nivel_original = fila["Resultado examen o curso"]
    return {'descripcion': nivel_original.capitalize(), "mce": None}
  return{'descripcion': '', 'mce': ''}


def formatear_longitud_nombre(nombre_str):
  normal_length = 35
  normal_fz = 36
  long_length = 36
  long_fz = 35
  very_long_fz = 34

  rt = RichText() # crea un objeto de texto enriquecido vacío
  longitud_nombre = len(nombre_str)

  if longitud_nombre <= normal_length:
    font_size = normal_fz
  elif longitud_nombre <= long_length:
    font_size = long_fz
  else:
    font_size = very_long_fz
  rt.add(
    nombre_str,
    font = 'Open Sans',
    size = font_size * 2,
    bold = True,
    color = "#808080"
  )
  return rt
#--------------------------------------------------------------------------


def generar_documento(fila):
  print(f"Generando documento para {fila['Nombres']}")
  try: 
    detalle_servicio = fila["Detalle de servicio"]
    idioma = fila["Idioma"]
    nombre_base_plantilla = detalle_servicio.replace(' ', '_')
    nombre_plantilla = f"{nombre_base_plantilla}_{idioma}.docx"
    #unir dos rutas 'os.path.join'
    ruta_plantilla = os.path.join(config.RUTA_PLANTILLAS, nombre_plantilla)
    if not os.path.exists(ruta_plantilla):
      print(f"❌ ¡ERROR! Plantilla no encontrada.")
      return None #detiene el proceso
    print(f"Se utilizó la plantilla {nombre_plantilla}")

    #Prepacióm de datos, llamada a las funciones especialistas
    contexto = {}

    if "Examen de comprensión de textos" in detalle_servicio:
      contexto['nombres_completos'] = str(fila['Nombres']).title()  
    else:
      nombre_original = str(fila['Nombres']).title()
      nombre_formateado = formatear_longitud_nombre(nombre_original)
      #print("\n--- DEBUG: Revisando el objeto RichText ---")
      #print(f"Tipo de objeto para el nombre: {type(nombre_formateado)}")
      #print(f"Contenido XML del objeto: {nombre_formateado.xml}")
      #print("-----------------------------------------\n")
      contexto['nombres_completos'] = nombre_formateado
    contexto['codigo_doc'] = str(fila['Código'])
    fecha_examen = formatear_fecha(fila['Fecha examen o curso'], idioma, detalle_servicio)
    fecha_emision = formatear_fecha(datetime.now(), idioma, detalle_servicio)
    nivel_info = procesar_nivel(fila, idioma, detalle_servicio)
    contexto.update({f'examen_{k}': v for k, v in fecha_examen.items()})
    contexto.update({f'emision_{k}': v for k, v in fecha_emision.items()})
    contexto['nivel_descripcion'] = nivel_info['descripcion']
    contexto['nivel_codigo'] = nivel_info['mce']
    if "Examen de comprensión de textos" in detalle_servicio:
      contexto['idioma_texto'] = fila['Idioma']
    else:
      contexto['idioma_texto'] = idioma
    print("     Contexto de datos ensamblado con éxito.")

    #Producción Final (renderizar, guardar, convertir)
    doc = DocxTemplate(ruta_plantilla)
    doc.render(contexto) #toma el diccionario 'contexto' y recorre el documento de word

    nombre_archivo = f"{fila['Código']} - {detalle_servicio} - {idioma} - {fila['Nombres']}" # el nombre del archivo final
    ruta_salida_docx = os.path.join(config.RUTA_SALIDAS, f"{nombre_archivo}.docx") #une la ruta hacia tu carpeta de salida
    ruta_salida_pdf = os.path.join(config.RUTA_SALIDAS, f"{nombre_archivo}.pdf")
    doc.save(ruta_salida_docx)
    print(f"Documento word temporal generando en: {ruta_salida_docx}")
    convert(ruta_salida_docx, ruta_salida_pdf)
    os.remove(ruta_salida_docx) # eliminar el documento word
    print(f"✅ Documento PDF final generado para {fila['Nombres']}.")
    return ruta_salida_pdf
  except Exception as e:
    print(f"❌ Error inesperado durante la generación del documento: {e}")
    return None
  


def crear_qr_firmar(fila, ruta_pdf):
  #construir el texto de validacióm para el qr
  nombre_qr = fila['Nombres']
  codigo_qr = fila['Código']
  correo_validacion = 'idiomas@oficinas-upc.pe'

  texto_qr = f"""
Idiomas Cayetano
---------------------
Estudiante: {nombre_qr}
Cod_documento: {str(codigo_qr)}
---------------------
Para corroborar la autenticidad de este documneto, por favor envíe un correo a:
{correo_validacion}
"""
  #Creación de la imagen QR
  qr_img = qrcode.make(texto_qr)
  ruta_qr_temp = os.path.join(config.RUTA_TEMP_QR, f"{fila['Código']}_qr.png")
  qr_img.save(ruta_qr_temp)
  print("imagen QR de validación creada")

  #Insertar qr en el pdf
  doc_pdf = fitz.open(ruta_pdf) #abre el documneto pdf que fue creaado en la fase anterior
  pagina = doc_pdf[0] #la primera página
  ancho_qr = 72 #1 inch
  pos_x = pagina.rect.width - ancho_qr - 36
  pos_y = pagina.rect.height - ancho_qr - 36
  rectangulo_qr = fitz.Rect(pos_x, pos_y, pos_x + ancho_qr, pos_y + ancho_qr)
  pagina.insert_image(rectangulo_qr, filename = ruta_qr_temp)

  # Guardado y limpieza
  doc_pdf.saveIncr()
  doc_pdf.close()
  os.remove(ruta_qr_temp)


