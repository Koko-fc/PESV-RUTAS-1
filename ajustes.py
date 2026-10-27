"""
Ajustes de despliegue de la app PESV.

Cada valor se busca, en este orden, en:
  1) variables de entorno
  2) "Secrets" de Streamlit (en la nube: panel de la app > Settings > Secrets;
     en tu PC: archivo .streamlit/secrets.toml)
  3) el valor por defecto de este archivo

Sin ningún secreto configurado la app funciona como siempre, 100% local.
"""
import os


def _leer(nombre, defecto=""):
    valor = os.environ.get(nombre)
    if valor is not None:
        return valor
    try:
        import streamlit as st
        if nombre in st.secrets:
            return st.secrets[nombre]
    except Exception:
        pass
    return defecto


def _como_bool(valor):
    return str(valor).strip().lower() in ("1", "true", "si", "sí", "yes", "verdadero")


# Si es True, la app SOLO muestra "Consultar ruta" (versión para conductores publicada en internet).
MODO_PUBLICO = _como_bool(_leer("MODO_PUBLICO", False))

# Registro de consultas en Google Sheets (vía Google Apps Script). Si faltan, se guarda en un archivo local.
SHEETS_URL = str(_leer("SHEETS_URL", "")).strip()
SHEETS_TOKEN = str(_leer("SHEETS_TOKEN", "")).strip()

# Colombia no usa horario de verano: UTC-5 fijo. Evita que en la nube (que trabaja en UTC) salgan horas corridas.
UTC_OFFSET_HORAS = int(float(_leer("UTC_OFFSET_HORAS", -5)))

# Texto que ven los conductores bajo el formulario. Ajústalo con la política de datos de la empresa.
AVISO_DATOS = str(_leer(
    "AVISO_DATOS",
    "Tu nombre y cédula se registran como evidencia de que consultaste la ruta antes de salir "
    "(Plan Estratégico de Seguridad Vial de DELTA GLOBAL S.A.S.).",
))
