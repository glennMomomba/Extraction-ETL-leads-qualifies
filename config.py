import os
from dotenv import load_dotenv

# Charger les variables du fichier .env
load_dotenv()

# Module-level defaults to avoid name-resolution issues inside the class body
THEME = os.getenv("THEME_RECHERCHE", "Agence immobilière")
# Récupère la chaîne de caractères et la découpe proprement par virgule
ZONES_RAW = os.getenv("ZONES", "Rabat")

class Config:
    THEME = THEME
    ZONES_RAW = ZONES_RAW
    SECTORS = [{"query": f"{THEME} {z.strip()}", "zone": z.strip()} for z in ZONES_RAW.split(",") if z.strip()]
    
    # Variables techniques sécurisées
    HEADLESS = os.getenv("HEADLESS_MODE", "True").lower() in ("true", "1", "yes")
    TIMEOUT = int(os.getenv("TIMEOUT_PAGE", 5000))
    MAX_LEADS = int(os.getenv("MAX_LEADS_PER_ZONE", 15))
    
    # Nom du fichier final dynamique basé sur le thème
    OUTPUT_FILENAME = f"Base_Prospects_{THEME.replace(' ', '_')}_DeepData.xlsx"