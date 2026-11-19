import os
import re
from urllib.parse import urljoin, urlparse
from playwright.sync_api import sync_playwright

URL = "https://gffsolutions.ups.com/IAF-Surcharges.html"
CARPETA_DESTINO = "descargas"
os.makedirs(CARPETA_DESTINO, exist_ok=True)


def limpiar_nombre(texto: str) -> str:
    limpio = re.sub(r"[^a-zA-Z0-9_-]+", "_", texto).strip("_")
    return limpio[:80] or "ups_iaf_surcharge"


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

        # Buscamos el primer enlace vinculado a IAF / International Air Freight Capacity Surcharge
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

        # Intentamos capturar descarga directa al hacer click; si abre en otra pestaña o responde directo, lo resolvemos
        try:
            with page.expect_download(timeout=10000) as download_info:
                target.click()
            download = download_info.value
            suggested = download.suggested_filename or f"{nombre_base}.pdf"
            ext = os.path.splitext(suggested)[1] or ".pdf"
            ruta_archivo = os.path.join(CARPETA_DESTINO, f"{nombre_base}{ext}")
            download.save_as(ruta_archivo)
        except Exception:
            # Si no dispara evento download, consultamos la URL directamente usando las cookies/sesión del contexto
            resp = context.request.get(target_url)
            if not resp.ok:
                raise RuntimeError(f"Error al obtener el archivo ({resp.status}) desde {target_url}")

            content_type = resp.headers.get("content-type", "").lower()
            ext = os.path.splitext(urlparse(target_url).path)[1].lower()
            if not ext:
                ext = ".pdf" if "pdf" in content_type else ".html"

            ruta_archivo = os.path.join(CARPETA_DESTINO, f"{nombre_base}{ext}")
            with open(ruta_archivo, "wb") as f:
                f.write(resp.body())

        print(f"Archivo guardado en: {ruta_archivo}")
        browser.close()


if __name__ == "__main__":
    main()
