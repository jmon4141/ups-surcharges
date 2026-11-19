import hashlib
import os
import re
from urllib.parse import urljoin, urlparse
from playwright.sync_api import sync_playwright

URL = "https://gffsolutions.ups.com/IAF-Surcharges.html"
CARPETA_DESTINO = "descargas"
ARCHIVO_HASH = os.path.join(CARPETA_DESTINO, "ultimo.sha256")
os.makedirs(CARPETA_DESTINO, exist_ok=True)


def limpiar_nombre(texto: str) -> str:
    limpio = re.sub(r"[^a-zA-Z0-9_-]+", "_", texto).strip("_")
    return limpio[:80] or "ups_iaf_surcharge"


def calcular_sha256(datos: bytes) -> str:
    return hashlib.sha256(datos).hexdigest()


def escribir_output(clave: str, valor: str):
    github_output = os.environ.get("GITHUB_OUTPUT")
    if github_output:
        with open(github_output, "a", encoding="utf-8") as f:
            f.write(f"{clave}={valor}\n")


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
            accept_downloads=True,
        )
        page = context.new_page()
        page.set_default_timeout(30000)

        page.goto(URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)

        locator = page.locator("a[href]").filter(
            has_text=re.compile(r"international air freight capacity surcharge|\biaf\b", re.I)
        )
        if locator.count() == 0:
            locator = page.locator("a[href*='IAF'], a[href*='iaf'], a[href*='.pdf']")
        if locator.count() == 0:
            raise RuntimeError("No se encontró el enlace esperado en la página.")

        target = locator.first
        texto_enlace = " ".join((target.inner_text() or "").split())
        href = target.get_attribute("href")
        target_url = urljoin(URL, href)
        nombre_base = limpiar_nombre(texto_enlace or urlparse(target_url).path)

        try:
            with page.expect_download(timeout=10000) as download_info:
                target.click()
            download = download_info.value
            suggested = download.suggested_filename or f"{nombre_base}.pdf"
            ext = os.path.splitext(suggested)[1] or ".pdf"
            temp_path = download.path()
            with open(temp_path, "rb") as f:
                contenido = f.read()
        except Exception:
            resp = context.request.get(target_url)
            if not resp.ok:
                raise RuntimeError(f"Error al obtener el archivo ({resp.status}) desde {target_url}")
            content_type = resp.headers.get("content-type", "").lower()
            ext = os.path.splitext(urlparse(target_url).path)[1].lower()
            if not ext:
                ext = ".pdf" if "pdf" in content_type else ".html"
            contenido = resp.body()

        browser.close()

    nuevo_hash = calcular_sha256(contenido)
    hash_anterior = ""
    if os.path.exists(ARCHIVO_HASH):
        with open(ARCHIVO_HASH, "r", encoding="utf-8") as f:
            hash_anterior = f.read().strip()

    if nuevo_hash == hash_anterior:
        print("El archivo no cambió desde la última revisión. No se envía correo.")
        escribir_output("cambio", "false")
        return

    ruta_archivo = os.path.join(CARPETA_DESTINO, f"{nombre_base}{ext}")
    with open(ruta_archivo, "wb") as f:
        f.write(contenido)

    with open(ARCHIVO_HASH, "w", encoding="utf-8") as f:
        f.write(nuevo_hash)

    print(f"Cambio detectado. Archivo guardado en: {ruta_archivo}")
    escribir_output("cambio", "true")
    escribir_output("archivo_path", ruta_archivo)
    escribir_output("archivo_nombre", f"{nombre_base}{ext}")


if __name__ == "__main__":
    main()
