"""
buscar_leads.py
Busca negocios en Montevideo sin sitio web usando Google Maps Places API.
Exporta resultados a leads.csv
"""

import os
import csv
import time
import googlemaps

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
# ───────────────────────────────────────────────────────────────────────────


def buscar_categoria(gmaps, categoria):
    """Devuelve lista de place_id para una categoría en la ciudad."""
    place_ids = []
    query = f"{categoria} {CIUDAD}"
    print(f"  Buscando: {query}")

    try:
        resp = gmaps.places(query=query)
    except Exception as e:
        print(f"  ERROR en búsqueda: {e}")
        return place_ids

    paginas = 0
    while resp and paginas < MAX_PAGES_PER_CATEGORIA:
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


def obtener_detalle(gmaps, place_id):
    """Devuelve detalle del negocio o None si falla."""
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
        print(f"  ERROR detalle {place_id}: {e}")
        return None


def main():
    if API_KEY == "TU_API_KEY_AQUI":
        print("ERROR: Configurá tu API key.")
        print("  export GOOGLE_MAPS_API_KEY='tu_key'")
        print("  o editá la variable API_KEY en este archivo.")
        return

    gmaps = googlemaps.Client(key=API_KEY)
    leads = []
    vistos = set()   # evitar duplicados entre categorías

    for categoria in CATEGORIAS:
        print(f"\n[{categoria.upper()}]")
        place_ids = buscar_categoria(gmaps, categoria)
        print(f"  Encontrados: {len(place_ids)} negocios")

        sin_web = 0
        for pid in place_ids:
            if pid in vistos:
                continue
            vistos.add(pid)

            time.sleep(DELAY_BETWEEN_REQUESTS)
            detalle = obtener_detalle(gmaps, pid)
            if not detalle:
                continue

            # Sin website = lead potencial
            if not detalle.get("website"):
                sin_web += 1
                leads.append({
                    "nombre":     detalle.get("name", ""),
                    "telefono":   detalle.get("formatted_phone_number", ""),
                    "direccion":  detalle.get("formatted_address", ""),
                    "categoria":  categoria,
                    "rating":     detalle.get("rating", ""),
                    "reseñas":    detalle.get("user_ratings_total", ""),
                    "website":    "",
                    "contactado": "",   # columna para marcar manualmente
                    "notas":      "",
                })

        print(f"  Sin web: {sin_web}")

    # ── EXPORTAR CSV ────────────────────────────────────────────────────────
    if not leads:
        print("\nNo se encontraron leads.")
        return

    fieldnames = ["nombre", "telefono", "direccion", "categoria",
                  "rating", "reseñas", "website", "contactado", "notas"]

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(leads)

    print(f"\n{'='*50}")
    print(f"TOTAL LEADS: {len(leads)}")
    print(f"Exportado a: {OUTPUT_FILE}")
    print(f"{'='*50}")


if __name__ == "__main__":
    main()
