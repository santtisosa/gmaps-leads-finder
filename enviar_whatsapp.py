"""
enviar_whatsapp.py
Envía mensajes de WhatsApp personalizados a cada lead en leads.csv.

Requisitos:
  - Chrome/Chromium instalado
  - WhatsApp Web abierto y logueado en el navegador por defecto
  - leads.csv generado por buscar_leads.py

Uso:
  venv/bin/python enviar_whatsapp.py
  venv/bin/python enviar_whatsapp.py --csv leads.csv --delay 30 --limite 20
"""

import csv
import time
import argparse
import re
import os
import openpyxl
import pywhatkit as kit
from dotenv import load_dotenv
from log_config import setup_logging

load_dotenv()

log = setup_logging("enviar_whatsapp")

# ── CONFIG ──────────────────────────────────────────────────────────────────
CSV_FILE   = "leads.xlsx"
DELAY_SEG  = 25
LIMITE     = None
WAIT_TIME  = 15
CLOSE_TIME = 3
# ────────────────────────────────────────────────────────────────────────────

MENSAJE_TEMPLATE = (
    "Hola, vi que {nombre} todavía no tiene sitio web. "
    "Te puedo armar uno en menos de 14 días, con dominio incluido y sin letra chica. "
    "Antes de cualquier pago te mando una maqueta gratis para que veas cómo quedaría. "
    "¿Te interesa? — Santiago | santiagososa.dev"
)


def limpiar_numero(telefono):
    """
    Convierte un teléfono a formato internacional sin + ni espacios.
    Asume Uruguay (+598) si no tiene código de país.
    """
    digits = re.sub(r"[^\d+]", "", telefono)

    if digits.startswith("+"):
        return digits
    elif digits.startswith("598"):
        return "+" + digits
    elif len(digits) >= 8:
        return "+598" + digits.lstrip("0")
    return None


def parse_args():
    parser = argparse.ArgumentParser(description="Envía mensajes WA a leads del CSV")
    parser.add_argument("--csv",    default=CSV_FILE,  help=f"Archivo CSV (default: {CSV_FILE})")
    parser.add_argument("--delay",  default=DELAY_SEG, type=int, help=f"Segundos entre mensajes (default: {DELAY_SEG})")
    parser.add_argument("--limite", default=LIMITE,    type=int, help="Máximo de mensajes a enviar (default: todos)")
    parser.add_argument("--dry-run", action="store_true", help="Muestra qué haría pero no envía nada")
    return parser.parse_args()


def cargar_leads(xlsx_path):
    """Lee el Excel y devuelve leads con teléfono no contactados."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    headers = [str(c.value).strip().lower() if c.value else "" for c in ws[1]]
    leads = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data = dict(zip(headers, row))
        if data.get("contactado", "") and str(data["contactado"]).strip():
            continue
        if not data.get("teléfono", "") and not data.get("telefono", ""):
            continue
        data["telefono"] = data.get("teléfono") or data.get("telefono") or ""
        data["nombre"]   = data.get("nombre", "")
        leads.append(data)
    return leads


def marcar_contactado(xlsx_path, nombre):
    """Marca un lead como contactado en el Excel."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    headers = [str(c.value).strip().lower() if c.value else "" for c in ws[1]]
    contactado_col = None
    nombre_col     = None
    for i, h in enumerate(headers, 1):
        if h in ("contactado",):   contactado_col = i
        if h == "nombre":          nombre_col = i
    if not contactado_col or not nombre_col:
        return
    for row in ws.iter_rows(min_row=2):
        if str(row[nombre_col - 1].value).strip() == nombre:
            row[contactado_col - 1].value = "✓"
            break
    wb.save(xlsx_path)
    log.debug(f"Marcado contactado: {nombre}")


def main():
    args = parse_args()

    log.info(f"Iniciando envío de mensajes — archivo: {args.csv}")

    if not os.path.exists(args.csv):
        log.error(f"No se encontró {args.csv}. Corré primero buscar_leads.py")
        return

    leads = cargar_leads(args.csv)
    log.info(f"Leads sin contactar: {len(leads)}")

    if not leads:
        log.info("No hay leads pendientes.")
        return

    if args.limite:
        leads = leads[:args.limite]
        log.info(f"Límite: {args.limite} mensajes")

    if args.dry_run:
        log.info("\n[DRY RUN — no se envía nada]\n")
        for lead in leads:
            numero = limpiar_numero(lead["telefono"])
            mensaje = MENSAJE_TEMPLATE.format(nombre=lead["nombre"])
            log.info(f"  → {lead['nombre']}")
            log.info(f"     Tel: {numero}")
            log.info(f"     Msg: {mensaje[:80]}...")
            log.info("")
        return

    log.info("\n⚠  Asegurate de tener WhatsApp Web abierto y logueado en Chrome.")
    input("   Presioná ENTER cuando estés listo...\n")

    enviados = 0
    errores  = 0

    for i, lead in enumerate(leads, 1):
        nombre  = lead["nombre"]
        numero  = limpiar_numero(lead["telefono"])
        mensaje = MENSAJE_TEMPLATE.format(nombre=nombre)

        if not numero:
            log.warning(f"[{i}] {nombre} — número inválido: {lead['telefono']}")
            errores += 1
            continue

        log.info(f"[{i}/{len(leads)}] Enviando a {nombre} ({numero})...")

        try:
            kit.sendwhatmsg_instantly(
                phone_no=numero,
                message=mensaje,
                wait_time=WAIT_TIME,
                tab_close=True,
                close_time=CLOSE_TIME,
            )
            marcar_contactado(args.csv, nombre)
            enviados += 1
            log.info(f"  ✓ Enviado")
            log.debug(f"  Mensaje enviado a {nombre} ({numero})")

        except Exception as e:
            log.error(f"  ✗ Error enviando a {nombre}: {e}")
            errores += 1

        if i < len(leads):
            log.info(f"  Esperando {args.delay}s...")
            time.sleep(args.delay)

    log.info(f"\n{'='*50}")
    log.info(f"Enviados:  {enviados}")
    log.info(f"Errores:   {errores}")
    log.info(f"CSV actualizado: columna 'contactado' marcada con ✓")
    log.info(f"{'='*50}")


if __name__ == "__main__":
    main()
