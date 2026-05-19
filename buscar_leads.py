"""
buscar_leads.py
Busca negocios en Montevideo sin sitio web usando Google Maps Places API.
Exporta resultados a leads.csv
"""

import os
import csv
import time
import argparse
import googlemaps
from dotenv import load_dotenv
from tqdm import tqdm

load_dotenv()  # carga .env automáticamente si existe

# ── CONFIG ─────────────────────────────────────────────────────────────────
API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "TU_API_KEY_AQUI")

CIUDAD = "Montevideo, Uruguay"

CATEGORIAS = [
    "peluquería",
    "barbería",
    "taller mecánico",
    "restaurante",
    "panadería",
    "ferretería",
    "veterinaria",
    "gimnasio",
    "lavandería",
    "cerrajería",
    "dentista",
    "kiosco",
    "carnicería",
    "frutería",
    "farmacia",
]

OUTPUT_FILE = "leads.csv"
DELAY_BETWEEN_REQUESTS = 0.2   # segundos entre llamadas API
MAX_PAGES_PER_CATEGORIA = 3    # cada página = 20 resultados → max 60 por categoría
MIN_RESENAS = 5                # mínimo de reseñas para considerar el negocio (filtra negocios fantasma)
# ───────────────────────────────────────────────────────────────────────────


def buscar_categoria(gmaps, categoria, ciudad=CIUDAD, max_paginas=MAX_PAGES_PER_CATEGORIA):
    """Devuelve lista de place_id para una categoría en la ciudad."""
    place_ids = []
    query = f"{categoria} {ciudad}"
    print(f"  Buscando: {query}")

    try:
        resp = gmaps.places(query=query)
    except Exception as e:
        print(f"  ERROR en búsqueda: {e}")
        return place_ids

    paginas = 0
    while resp and paginas < max_paginas:
        for place in resp.get("results", []):
            place_ids.append(place["place_id"])

        paginas += 1
        next_token = resp.get("next_page_token")
        if not next_token:
            break

        time.sleep(2)  # Google requiere ~2s antes de usar next_page_token
        try:
            resp = gmaps.places(query=query, page_token=next_token)
        except Exception as e:
            print(f"  ERROR paginación: {e}")
            break

    return place_ids


def obtener_detalle(gmaps, place_id, reintentos=3):
    """Devuelve detalle del negocio. Reintenta con backoff exponencial."""
    for intento in range(reintentos):
        try:
            result = gmaps.place(
                place_id,
                fields=[
                    "name",
                    "website",
                    "formatted_phone_number",
                    "formatted_address",
                    "rating",
                    "user_ratings_total",
                ],
            )
            return result.get("result", {})
        except Exception as e:
            if intento < reintentos - 1:
                espera = 2 ** intento   # 1s, 2s, 4s
                time.sleep(espera)
            else:
                tqdm.write(f"  ERROR detalle {place_id}: {e}")
    return None


def parse_args():
    parser = argparse.ArgumentParser(description="Busca negocios sin web en Google Maps")
    parser.add_argument("--ciudad",   default=CIUDAD,              help="Ciudad a buscar (default: Montevideo, Uruguay)")
    parser.add_argument("--output",   default=OUTPUT_FILE,          help="Archivo CSV de salida (default: leads.csv)")
    parser.add_argument("--paginas",  default=MAX_PAGES_PER_CATEGORIA, type=int, help="Páginas por categoría (default: 3)")
    parser.add_argument("--categorias", nargs="+", default=CATEGORIAS, help="Categorías a buscar")
    parser.add_argument("--min-resenas", default=MIN_RESENAS, type=int, help="Mínimo de reseñas para incluir un negocio (default: 5)")
    return parser.parse_args()


def main():
    args = parse_args()

    if API_KEY == "TU_API_KEY_AQUI":
        print("ERROR: Configurá tu API key.")
        print("  export GOOGLE_MAPS_API_KEY='tu_key'")
        print("  o creá un archivo .env con GOOGLE_MAPS_API_KEY=tu_key")
        return

    gmaps = googlemaps.Client(key=API_KEY)
    leads = []
    vistos = set()

    # Cargar nombres ya guardados para no duplicar entre corridas
    nombres_existentes = set()
    if os.path.exists(args.output):
        with open(args.output, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                nombres_existentes.add(row.get("nombre", "").strip().lower())
        print(f"Leads existentes en {args.output}: {len(nombres_existentes)} (se saltarán duplicados)")

    print(f"Ciudad: {args.ciudad}")
    print(f"Categorías: {len(args.categorias)}")
    print(f"Output: {args.output}\n")

    for categoria in args.categorias:
        print(f"\n[{categoria.upper()}]")
        place_ids = buscar_categoria(gmaps, categoria, ciudad=args.ciudad, max_paginas=args.paginas)
        print(f"  Encontrados: {len(place_ids)} negocios")

        sin_web = 0
        nuevos = [p for p in place_ids if p not in vistos]
        for pid in tqdm(nuevos, desc=f"  {categoria}", unit="neg", leave=False):
            vistos.add(pid)

            time.sleep(DELAY_BETWEEN_REQUESTS)
            detalle = obtener_detalle(gmaps, pid)
            if not detalle:
                continue

            # Sin website = lead potencial
            nombre = detalle.get("name", "")
            resenas = detalle.get("user_ratings_total") or 0
            if not detalle.get("website") and resenas >= args.min_resenas and nombre.strip().lower() not in nombres_existentes:
                sin_web += 1
                nombres_existentes.add(nombre.strip().lower())
                leads.append({
                    "nombre":     nombre,
                    "telefono":   detalle.get("formatted_phone_number", ""),
                    "direccion":  detalle.get("formatted_address", ""),
                    "categoria":  categoria,
                    "rating":     detalle.get("rating", ""),
                    "reseñas":    detalle.get("user_ratings_total", ""),
                    "website":    "",
                    "contactado": "",
                    "notas":      "",
                })

        print(f"  Sin web: {sin_web}")

    # ── EXPORTAR CSV ────────────────────────────────────────────────────────
    if not leads:
        print("\nNo se encontraron leads.")
        return

    fieldnames = ["nombre", "telefono", "direccion", "categoria",
                  "rating", "reseñas", "website", "contactado", "notas"]

    with open(args.output, "a", newline="", encoding="utf-8") as f:
        # "a" = append: si el archivo existe agrega, si no lo crea
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if f.tell() == 0:
            writer.writeheader()   # solo escribe header si el archivo está vacío
        writer.writerows(leads)

    print(f"\n{'='*50}")
    print(f"TOTAL LEADS: {len(leads)}")
    print(f"Exportado a: {args.output}")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
