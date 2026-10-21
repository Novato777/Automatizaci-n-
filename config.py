import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"

def load_env():
    """Carga variables desde el archivo .env sin dependencias externas."""
    env_vars = {}
    if ENV_PATH.exists():
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    env_vars[key.strip()] = val.strip().strip('"').strip("'")
    return env_vars

env = load_env()

GMAIL_USER = os.getenv("GMAIL_USER", env.get("GMAIL_USER", "cardozobrayan334@gmail.com"))
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", env.get("GMAIL_APP_PASSWORD", "")).replace(" ", "")
SENDER_NAME = os.getenv("SENDER_NAME", env.get("SENDER_NAME", "NeXo Software"))
EXCEL_PATH = BASE_DIR / env.get("EXCEL_PATH", "NeXo_Gestion_Prospectos_y_Correos.xlsx")

try:
    DELAY_MIN = int(env.get("DELAY_MIN_SECONDS", 45))
    DELAY_MAX = int(env.get("DELAY_MAX_SECONDS", 75))
except ValueError:
    DELAY_MIN = 45
    DELAY_MAX = 75
