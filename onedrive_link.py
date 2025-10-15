import os
import json
import urllib.parse
from typing import Optional

import requests
try:
  import msal  # type: ignore
except ImportError:
  msal = None  # el usuario debe instalar msal


def _posibles_raices_onedrive():
  # Variables típicas de OneDrive en Windows
  rutas = []
  for var in ["OneDrive", "OneDriveCommercial", "OneDriveConsumer"]:
    ruta = os.environ.get(var)
    if ruta and os.path.isdir(ruta):
      rutas.append(ruta)
  # Carpeta "OneDrive - <Org>" en el perfil de usuario
  usuario = os.path.expanduser("~")
  if os.path.isdir(usuario):
    for nombre in os.listdir(usuario):
      if nombre.startswith("OneDrive"):
        candidato = os.path.join(usuario, nombre)
        if os.path.isdir(candidato):
          rutas.append(candidato)
  # Eliminar duplicados preservando orden
  vistas = []
  res = []
  for r in rutas:
    if r.lower() not in vistas:
      vistas.append(r.lower())
      res.append(r)
  return res


def _relativo_a_onedrive(local_path: str) -> Optional[tuple[str, str]]:
  """
  Devuelve (raiz, relativo) si local_path está dentro de alguna raíz de OneDrive.
  """
  lp = os.path.abspath(local_path)
  for raiz in _posibles_raices_onedrive():
    raiz_abs = os.path.abspath(raiz)
    try:
      rel = os.path.relpath(lp, raiz_abs)
    except ValueError:
      continue
    if not rel.startswith(".."):
      return raiz_abs, rel.replace("\\", "/")
  return None


def _obtener_token_msal() -> Optional[str]:
  """
  Usa MSAL Device Code Flow para obtener un token de Microsoft Graph.
  Requiere variables de entorno:
    - AZURE_CLIENT_ID (app registrada con permisos Delegated Files.Read.All o Files.ReadWrite.All)
    - AZURE_TENANT_ID (o 'common')
  """
  if msal is None:
    print("⚠️ Debe instalar 'msal' para generar vínculos de OneDrive: pip install msal")
    return None

  client_id = os.environ.get("AZURE_CLIENT_ID")
  tenant = os.environ.get("AZURE_TENANT_ID", "common")
  if not client_id:
    print("⚠️ Configure la variable de entorno AZURE_CLIENT_ID (aplicación Azure AD registrada)")
    return None

  app = msal.PublicClientApplication(client_id=client_id, authority=f"https://login.microsoftonline.com/{tenant}")
  scopes = ["Files.ReadWrite.All", "offline_access"]

  # Intentar cache silencioso primero (si ya se autenticó antes)
  cuentas = app.get_accounts()
  if cuentas:
    result = app.acquire_token_silent(scopes, account=cuentas[0])
    if result and "access_token" in result:
      return result["access_token"]

  # Device code flow interactivo en terminal (una sola vez)
  flow = app.initiate_device_flow(scopes=scopes)
  if "user_code" not in flow:
    print("❌ No se pudo iniciar el device code flow para MSAL")
    return None
  print(flow["message"])  # Instrucciones para autenticar una sola vez
  result = app.acquire_token_by_device_flow(flow)
  if "access_token" in result:
    return result["access_token"]
  print(f"❌ Error al obtener token: {result}")
  return None


def crear_link_compartir_onedrive(local_path: str) -> Optional[str]:
  """
  Genera y devuelve un vínculo de uso anónimo ("Copiar vínculo") al archivo en OneDrive.
  - Debe estar dentro de una carpeta sincronizada de OneDrive.
  - Requiere Microsoft Graph (token MSAL).
  """
  rel = _relativo_a_onedrive(local_path)
  if not rel:
    print("⚠️ El archivo no está dentro de una carpeta sincronizada de OneDrive.")
    return None
  _, relativo = rel

  token = _obtener_token_msal()
  if not token:
    return None

  # Endpoint: POST /me/drive/root:{item-path}:/createLink
  encoded_path = urllib.parse.quote(relativo)
  url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_path}:/createLink"
  body = {"type": "view", "scope": "anonymous"}
  headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
  resp = requests.post(url, headers=headers, data=json.dumps(body), timeout=30)
  if resp.status_code in (200, 201):
    data = resp.json()
    link = data.get("link", {}).get("webUrl")
    if link:
      return link
    print("⚠️ Respuesta sin 'webUrl':", data)
    return None
  else:
    try:
      detalle = resp.json()
    except Exception:
      detalle = resp.text
    print(f"❌ Error Graph {resp.status_code}: {detalle}")
    return None