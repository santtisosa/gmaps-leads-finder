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
import tempfile
import re
import os
import pywhatkit as kit
from dotenv import load_dotenv

load_dotenv()

# ── CONFIG ──────────────────────────────────────────────────────────────────
CSV_FILE   = "leads.csv"
DELAY_SEG  = 25     # segundos entre mensajes (no spamear)
LIMITE     = None   # None = todos, o un número para probar (ej: 5)
WAIT_TIME  = 15     # segundos que espera antes de enviar (para que cargue WA Web)
CLOSE_TIME = 3      # segundos antes de cerrar la pestaña
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
    # Quitar todo excepto dígitos y +
    digits = re.sub(r"[^\d+]", "", telefono)

    if digits.startswith("+"):
        return digits   # ya tiene código de país
    elif digits.startswith("598"):
        return "+" + digits
    elif len(digits) >= 8:
        return "+598" + digits.lstrip("0")  # agrega Uruguay
    return None


def parse_args():
    parser = argparse.ArgumentParser(description="Envía mensajes WA a leads del CSV")
    parser.add_argument("--csv",    default=CSV_FILE, help=f"Archivo CSV (default: {CSV_FILE})")
    parser.add_argument("--delay",  default=DELAY_SEG, type=int, help=f"Segundos entre mensajes (default: {DELAY_SEG})")
    parser.add_argument("--limite", default=LIMITE, type=int, help="Máximo de mensajes a enviar (default: todos)")
    parser.add_argument("--dry-run", action="store_true", help="Muestra qué haría pero no envía nada")
    return parser.parse_args()


def cargar_leads(csv_path):
    """Lee el CSV y devuelve leads con teléfono no contactados."""
    leads = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("contactado", "").strip():
                continue   # ya fue contactado
            if not row.get("telefono", "").strip():
                continue   # sin teléfono
            leads.append(row)
    return leads


def marcar_contactado(csv_path, nombre):
    """Marca un lead como contactado en el CSV."""
    rows = []
    fieldnames = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        for row in reader:
            if row["nombre"] == nombre:
                row["contactado"] = "✓"
            rows.append(row)

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main():
    args = parse_args()

    if not os.path.exists(args.csv):
        print(f"ERROR: No se encontró {args.csv}. Corré primero buscar_leads.py")
        return

    leads = cargar_leads(args.csv)
    print(f"Leads sin contactar: {len(leads)}")

    if not leads:
        print("No hay leads pendientes.")
        return

    if args.limite:
        leads = leads[:args.limite]
        print(f"Límite: {args.limite} mensajes")

    if args.dry_run:
        print("\n[DRY RUN — no se envía nada]\n")
        for lead in leads:
            numero = limpiar_numero(lead["telefono"])
            mensaje = MENSAJE_TEMPLATE.format(nombre=lead["nombre"])
            print(f"  → {lead['nombre']}")
            print(f"     Tel: {numero}")
            print(f"     Msg: {mensaje[:80]}...")
            print()
        return

    print("\n⚠  Asegurate de tener WhatsApp Web abierto y logueado en Chrome.")
    input("   Presioná ENTER cuando estés listo...\n")

    enviados = 0
    errores  = 0

    for i, lead in enumerate(leads, 1):
        nombre  = lead["nombre"]
        numero  = limpiar_numero(lead["telefono"])
        mensaje = MENSAJE_TEMPLATE.format(nombre=nombre)

        if not numero:
            print(f"[{i}] ⚠  {nombre} — número inválido: {lead['telefono']}")
            errores += 1
            continue

        print(f"[{i}/{len(leads)}] Enviando a {nombre} ({numero})...")

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
            print(f"  ✓ Enviado")

        except Exception as e:
            print(f"  ✗ Error: {e}")
            errores += 1

        if i < len(leads):
            print(f"  Esperando {args.delay}s...")
            time.sleep(args.delay)

    print(f"\n{'='*50}")
    print(f"Enviados:  {enviados}")
    print(f"Errores:   {errores}")
    print(f"CSV actualizado: columna 'contactado' marcada con ✓")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
