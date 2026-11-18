import os
import re
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

URL = "https://ups.com"
CARPETA_DESTINO = "descargas"
os.makedirs(CARPETA_DESTINO, exist_ok=True)

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

resp = requests.get(URL, headers=headers, timeout=30)
resp.raise_for_status()

soup = BeautifulSoup(resp.text, "html.parser")
enlace = soup.find("a", string=re.compile(r"View International Air Freight Capacity Surcharge\s*-", re.I))

if not enlace or not enlace.get("href"):
    raise Exception("No se encontró el enlace del PDF en la página.")

pdf_url = urljoin(URL, enlace["href"])
texto = enlace.get_text(strip=True)
nombre_limpio = re.sub(r"[^a-zA-Z0-9_-]", "_", texto) + ".pdf"
ruta_archivo = os.path.join(CARPETA_DESTINO, nombre_limpio)

pdf_resp = requests.get(pdf_url, headers=headers, timeout=30)
pdf_resp.raise_for_status()

with open(ruta_archivo, "wb") as f:
    f.write(pdf_resp.content)

print(f"Descargado correctamente: {ruta_archivo}")
