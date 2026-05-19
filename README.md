# Google Maps Leads Finder

Script that searches Google Maps for local businesses **without a website** in Montevideo, Uruguay and exports them to a CSV file.

## What it does

- Searches 15+ business categories (hair salons, mechanics, restaurants, etc.)
- Checks each business for the presence of a website
- Exports leads (businesses without a website) to `leads.csv`

## Output

`leads.csv` includes:
| Field | Description |
|-------|-------------|
| nombre | Business name |
| telefono | Phone number |
| direccion | Address |
| categoria | Search category |
| rating | Google rating |
| reseñas | Number of reviews |
| contactado | Mark manually when contacted |
| notas | Your notes |

## Setup

**1. Clone the repo**
```bash
git clone https://github.com/your-username/gmaps-leads-finder.git
cd gmaps-leads-finder
```

**2. Create a virtual environment**
```bash
python3 -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

**3. Get a Google Maps API key**
- Go to [console.cloud.google.com](https://console.cloud.google.com)
- Create a project → enable **Places API** → create an API key
- Restrict the key to Places API only

**4. Set your API key**
```bash
cp .env.example .env
# Edit .env and paste your API key
```

Or export it directly:
```bash
export GOOGLE_MAPS_API_KEY="your_key_here"
```

**5. Run**
```bash
python buscar_leads.py
```

## Cost

Google Maps Platform gives **$200 free credit per month**. This script costs ~$10–15 per full run — well within the free tier.

## Customize

Edit `buscar_leads.py` to change:
- `CIUDAD` — target city
- `CATEGORIAS` — business types to search
- `MAX_PAGES_PER_CATEGORIA` — results per category (20 per page)

## License

MIT
