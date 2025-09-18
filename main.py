import pandas as pd
import config 

#--- FASE 1: LEER ARCHIVO EXCEL ---#
df = pd.read_excel(
  config.RUTA_EXCEL,
  sheet_name="Cert-Cons",
  skiprows=3,
  usecols="B:X"
)
pendiente_envio = df[
  (~df['Código'].isin(config.EXCLUIR_VALORES)) & (df['Estado'].isna())
].dropna(subset=['Año'])

if pendiente_envio.empty:
  print("👍No se encontraron nuevas solicitudes pendientes")
else:
  print(f"🔥Se encontraton {len(pendiente_envio)} solicitud(es) por enviar")


