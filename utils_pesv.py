"""
Utilidades para la app PESV:
- Conversión de link de Google My Maps a formato "embed" (iframe)
- Generación de PDF de una ficha de ruta
- Persistencia del registro de consultas (evidencia PESV)
"""
import json
import re
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from io import BytesIO

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, ListFlowable, ListItem,
    Image, KeepTogether,
)
from PIL import Image as PILImage, ImageOps
import requests

import ajustes

DATA_DIR = Path(__file__).parent / "data"
CONSULTAS_PATH = DATA_DIR / "consultas.json"
MAPAS_DIR = DATA_DIR / "mapas"      # capturas del mapa de cada ruta (se incluyen en el PDF)
MAX_ANCHO_PX = 1600                 # se reduce el tamaño de capturas muy grandes

NIVEL_COLOR_PDF = {
    "Alto": colors.HexColor("#F8CBAD"),
    "Medio": colors.HexColor("#FFE699"),
    "Bajo": colors.HexColor("#C6E0B4"),
}


# ---------------------------------------------------------------------------
# Google My Maps: edit link -> embed link
# ---------------------------------------------------------------------------
def convertir_a_embed(url: str) -> str:
    """
    Convierte un link de Google My Maps del tipo:
      https://www.google.com/maps/d/edit?mid=XXXX&usp=sharing
      https://www.google.com/maps/d/u/0/edit?mid=XXXX&usp=sharing
    en su versión embebible (iframe):
      https://www.google.com/maps/d/u/0/embed?mid=XXXX
    Si la url no matchea el patrón esperado, la devuelve tal cual.
    """
    if not url:
        return ""
    match = re.search(r"mid=([a-zA-Z0-9_-]+)", url)
    if not match:
        return url
    mid = match.group(1)
    return f"https://www.google.com/maps/d/u/0/embed?mid={mid}"


# ---------------------------------------------------------------------------
# Captura del mapa por ruta
# ---------------------------------------------------------------------------
def ruta_imagen_path(ruta):
    """Devuelve la Path de la captura del mapa de la ruta, o None si no hay / no existe."""
    nombre = (ruta or {}).get("mapa_imagen", "")
    if not nombre:
        return None
    p = MAPAS_DIR / Path(nombre).name
    return p if p.exists() else None


def guardar_imagen_mapa(data: bytes) -> str:
    """
    Guarda la captura subida (PNG/JPG) en data/mapas/ y devuelve el nombre del archivo.
    Lanza una excepción si el archivo no es una imagen válida.
    """
    MAPAS_DIR.mkdir(parents=True, exist_ok=True)
    img = PILImage.open(BytesIO(data))
    img = ImageOps.exif_transpose(img)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        fondo = PILImage.new("RGB", img.size, (255, 255, 255))
        fondo.paste(img, mask=img.split()[-1])
        img = fondo
    else:
        img = img.convert("RGB")
    if img.width > MAX_ANCHO_PX:
        alto = round(img.height * MAX_ANCHO_PX / img.width)
        img = img.resize((MAX_ANCHO_PX, alto), PILImage.LANCZOS)
    nombre = f"mapa_{uuid.uuid4().hex[:12]}.jpg"
    img.save(MAPAS_DIR / nombre, "JPEG", quality=88, optimize=True)
    return nombre


def eliminar_imagen_mapa(nombre):
    if not nombre:
        return
    try:
        (MAPAS_DIR / Path(nombre).name).unlink(missing_ok=True)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Registro de consultas
# ---------------------------------------------------------------------------
PENDIENTES_PATH = DATA_DIR / "consultas_pendientes.json"
TIMEOUT_REMOTO = 15


def ahora_colombia():
    """Hora actual de Colombia (UTC-5), aunque el servidor esté en otra zona horaria."""
    return datetime.now(timezone(timedelta(hours=ajustes.UTC_OFFSET_HORAS)))


def _leer_json(path):
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return []
    return []


def _escribir_json(path, datos):
    DATA_DIR.mkdir(exist_ok=True)
    tmp = path.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def remoto_configurado():
    return bool(ajustes.SHEETS_URL and ajustes.SHEETS_TOKEN)


def _llamar_remoto(accion, **datos):
    """Llama al servicio de Google Apps Script. Lanza excepción si algo falla."""
    payload = {"token": ajustes.SHEETS_TOKEN, "accion": accion, **datos}
    r = requests.post(
        ajustes.SHEETS_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "text/plain; charset=utf-8"},
        timeout=TIMEOUT_REMOTO,
    )
    r.raise_for_status()
    try:
        resp = r.json()
    except ValueError:
        raise RuntimeError("respuesta inesperada del servicio de registro")
    if not resp.get("ok"):
        raise RuntimeError(resp.get("error", "error desconocido"))
    return resp


# --- API usada por la app -------------------------------------------------
def cargar_consultas_estado():
    """Devuelve (lista_de_consultas, aviso). 'aviso' es None cuando todo está bien."""
    if not remoto_configurado():
        return _leer_json(CONSULTAS_PATH), None
    pendientes = _leer_json(PENDIENTES_PATH)
    try:
        remotas = _llamar_remoto("listar").get("registros", [])
    except Exception as e:
        return pendientes, (
            "No se pudo leer el registro en Google Sheets "
            f"({type(e).__name__}). Solo se muestran las consultas pendientes de enviar."
        )
    ids = {r.get("id") for r in remotas}
    extra = [p for p in pendientes if p.get("id") not in ids]
    aviso = (f"Hay {len(extra)} consulta(s) guardadas en este equipo que aún no llegan a Google Sheets; "
             "se enviarán con la próxima consulta.") if extra else None
    return remotas + extra, aviso


def cargar_consultas():
    return cargar_consultas_estado()[0]


def registrar_consulta(conductor: str, ruta_nombre: str, cedula: str = ""):
    """
    Registra la consulta. Devuelve:
      "local"     -> guardada en data/consultas.json (modo sin Google Sheets)
      "remoto"    -> guardada en Google Sheets
      "pendiente" -> Google Sheets no respondió; quedó en cola y se reenvía en la próxima consulta
    """
    registro = {
        "id": uuid.uuid4().hex,
        "fecha_hora": ahora_colombia().strftime("%Y-%m-%d %H:%M:%S"),
        "conductor": conductor.strip(),
        "cedula": cedula.strip(),
        "ruta": ruta_nombre,
    }
    if not remoto_configurado():
        consultas = _leer_json(CONSULTAS_PATH)
        consultas.append(registro)
        _escribir_json(CONSULTAS_PATH, consultas)
        return "local"

    pendientes = _leer_json(PENDIENTES_PATH)
    try:
        # Se envían juntas las pendientes y la nueva. El 'id' evita duplicados si se reintenta.
        _llamar_remoto("registrar", registros=pendientes + [registro])
        if pendientes:
            _escribir_json(PENDIENTES_PATH, [])
        return "remoto"
    except Exception:
        _escribir_json(PENDIENTES_PATH, pendientes + [registro])
        return "pendiente"


# ---------------------------------------------------------------------------
# PDF de la ficha
# ---------------------------------------------------------------------------
def generar_pdf_ficha(ruta: dict, conductor: str = "") -> BytesIO:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
        leftMargin=2 * cm, rightMargin=2 * cm,
    )
    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle(
        "TituloRuta", parent=styles["Title"], fontSize=16,
        textColor=colors.HexColor("#1F4E78"), spaceAfter=4,
    )
    subtitulo_style = ParagraphStyle(
        "Subtitulo", parent=styles["Normal"], fontSize=9,
        textColor=colors.HexColor("#555555"), spaceAfter=14,
    )
    campo_style = styles["Normal"]
    seccion_style = ParagraphStyle(
        "Seccion", parent=styles["Heading2"], fontSize=12,
        textColor=colors.white, backColor=colors.HexColor("#2E75B6"),
        spaceBefore=10, spaceAfter=6, leftIndent=4, borderPadding=4,
    )

    story = []
    story.append(Paragraph("Ficha de Ruta Frecuente — PESV", titulo_style))
    generado = f"Generado: {ahora_colombia().strftime('%Y-%m-%d %H:%M')}"
    if conductor:
        generado += f"  |  Consultado por: {conductor}"
    story.append(Paragraph(generado, subtitulo_style))

    story.append(Paragraph(ruta.get("nombre", ""), styles["Heading1"]))
    story.append(Spacer(1, 6))

    campos = [
        ("Cliente / Sede", ruta.get("cliente", "")),
        ("Origen", ruta.get("origen", "")),
        ("Destino", ruta.get("destino", "")),
        ("Frecuencia", ruta.get("frecuencia", "")),
        ("Recorrido", ruta.get("recorrido", "")),
        ("Distancia aproximada", ruta.get("distancia", "")),
        ("Tiempo estimado", ruta.get("tiempo", "")),
        ("Riesgos generales", ruta.get("riesgos", "")),
        ("Controles", ruta.get("controles", "")),
    ]
    tabla_data = [[Paragraph(f"<b>{k}</b>", campo_style), Paragraph(str(v or "—"), campo_style)]
                  for k, v in campos if v]
    if tabla_data:
        t = Table(tabla_data, colWidths=[4.5 * cm, 11 * cm])
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F2F2")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t)

    img_path = ruta_imagen_path(ruta)
    if img_path:
        iw, ih = ImageReader(str(img_path)).getSize()
        max_w, max_h = 15.5 * cm, 9.5 * cm
        escala = min(max_w / iw, max_h / ih)
        imagen = Image(str(img_path), width=iw * escala, height=ih * escala)
        marco = Table([[imagen]], colWidths=[iw * escala + 6])
        marco.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#BFBFBF")),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(KeepTogether([
            Paragraph("Mapa de la ruta", seccion_style),
            marco,
        ]))

    if ruta.get("puntos_criticos"):
        story.append(Paragraph("Puntos críticos", seccion_style))
        cabecera = [Paragraph("<b>Punto</b>", campo_style),
                    Paragraph("<b>Riesgo</b>", campo_style),
                    Paragraph("<b>Nivel</b>", campo_style)]
        filas = [cabecera]
        estilos_extra = []
        for i, pc in enumerate(ruta["puntos_criticos"], start=1):
            filas.append([
                Paragraph(pc.get("punto", ""), campo_style),
                Paragraph(pc.get("riesgo", ""), campo_style),
                Paragraph(pc.get("nivel", ""), campo_style),
            ])
            color = NIVEL_COLOR_PDF.get(pc.get("nivel", ""))
            if color:
                estilos_extra.append(("BACKGROUND", (2, i), (2, i), color))
        t2 = Table(filas, colWidths=[5 * cm, 8 * cm, 2.5 * cm])
        t2.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E75B6")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            *estilos_extra,
        ]))
        story.append(t2)

    if ruta.get("medidas_preventivas"):
        story.append(Paragraph("Medidas preventivas", seccion_style))
        items = [l.strip() for l in ruta["medidas_preventivas"].split("\n") if l.strip()]
        lista = ListFlowable(
            [ListItem(Paragraph(item, campo_style)) for item in items],
            bulletType="bullet",
        )
        story.append(lista)

    if ruta.get("mapa_url"):
        story.append(Spacer(1, 10))
        story.append(Paragraph("Mapa interactivo en línea (Google My Maps):", campo_style))
        story.append(Paragraph(f'<link href="{ruta["mapa_url"]}">{ruta["mapa_url"]}</link>', campo_style))

    story.append(Spacer(1, 16))
    story.append(Paragraph(
        "Este documento certifica que la ruta y sus riesgos fueron consultados "
        "previamente al desplazamiento, como parte del Plan Estratégico de "
        "Seguridad Vial (PESV).",
        subtitulo_style,
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer
