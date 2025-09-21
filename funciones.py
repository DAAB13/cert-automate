import config

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
  if "Examen de comprension de textos" in detalle_servicio:
    mes_es = fecha_x.strftime("%B")
    return {'texto_simple': f"el {fecha_x.day} de {mes_es} de {fecha_x.year}"}
  
  elif idioma == "Portugués":
    mes_es = fecha_x.strftime("%B")
    mes_portugues = config.MES_PORTUGUES.get(mes_es.lower(), mes_es)
    return {'texto_simple': f"em {fecha_x.day} de {mes_portugues} de {fecha_x.year}"}

  elif idioma == "ingles":
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
    mes_es = fecha_x.strftime("%B")
    return f"el {fecha_x.day} de {mes_es} de {fecha_x.year}"
