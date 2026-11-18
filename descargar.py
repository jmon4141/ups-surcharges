import os
import re
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright

URL = "https://ups.com"
CARPETA_DESTINO = "descargas"
os.makedirs(CARPETA_DESTINO, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    )
    page = context.new_page()
    page.goto(URL, wait_until="networkidle", timeout=60000)

    enlace = page.locator("a", has_text=re.compile(r"View International Air Freight Capacity Surcharge\s*-", re.I)).first
    href = enlace.get_attribute("href")
    texto = enlace.inner_text().strip()

    if not href:
        raise Exception("No se encontró el enlace del archivo.")

    pdf_url = urljoin(URL, href)
    nombre_limpio = re.sub(r"[^a-zA-Z0-9_-]", "_", texto) + ".pdf"
    ruta_archivo = os.path.join(CARPETA_DESTINO, nombre_limpio)

    response = context.request.get(pdf_url)
    if response.ok:
        with open(ruta_archivo, "wb") as f:
            f.write(response.body())
        print(f"Descargado correctamente: {ruta_archivo}")
    else:
        raise Exception(f"Error al descargar el PDF. Código HTTP: {response.status}")

    browser.close()
