import hashlib
import os
import re
from urllib.parse import urljoin, urlparse

import pandas as pd
import pdfplumber
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


def limpiar_celda(val) -> str:
    if val is None:
        return ""
    return " ".join(str(val).replace("\n", " ").split())


def parsear_valor(val: str):
    if not val:
        return None
    limpio = val.replace("$", "").replace(",", ".").strip()
    if limpio in ("-", "—", "N/A", "n/a"):
        return 0.0
    try:
        return float(limpio)
    except ValueError:
        return val


def _extraer_por_coordenadas(pdf) -> list[list[str]]:
    filas_reconstruidas = []
    for page in pdf.pages:
        palabras = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        if not palabras:
            continue

        palabras_ordenadas = sorted(palabras, key=lambda w: (w["top"], w["x0"]))
        lineas = []
        for w in palabras_ordenadas:
            if not lineas or abs(w["top"] - lineas[-1][-1]["top"]) > 4.0:
                lineas.append([w])
            else:
                lineas[-1].append(w)

        for linea in lineas:
            linea_sorted = sorted(linea, key=lambda w: w["x0"])
            celdas = []
            actual = [linea_sorted[0]["text"]]
            prev_x1 = linea_sorted[0]["x1"]
            for w in linea_sorted[1:]:
                if (w["x0"] - prev_x1) > 14.0:
                    celdas.append(" ".join(actual))
                    actual = [w["text"]]
                else:
                    actual.append(w["text"])
                prev_x1 = max(prev_x1, w["x1"])
            celdas.append(" ".join(actual))
            if len(celdas) >= 3:
                filas_reconstruidas.append(celdas)

    return filas_reconstruidas


def pdf_a_excel(ruta_pdf: str, ruta_excel: str):
    tablas_extraidas = []
    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            tablas = page.extract_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 4,
                    "join_tolerance": 4,
                }
            )
            if not tablas:
                tablas = page.extract_tables(
                    table_settings={
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                    }
                )
            if tablas:
                tablas_extraidas.extend(tablas)

        if not tablas_extraidas:
            coord_rows = _extraer_por_coordenadas(pdf)
            if coord_rows:
                max_cols = max(len(r) for r in coord_rows)
                tablas_extraidas.append([r + [""] * (max_cols - len(r)) for r in coord_rows])

    if not tablas_extraidas:
        with pdfplumber.open(ruta_pdf) as pdf:
            texto_paginas = [
                {"Pagina": i + 1, "Texto": p.extract_text() or ""}
                for i, p in enumerate(pdf.pages)
            ]
        with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
            pd.DataFrame(texto_paginas).to_excel(writer, sheet_name="Texto_Crudo", index=False)
        return

    tabla_principal = max(tablas_extraidas, key=lambda t: len(t) * (len(t[0]) if t else 0))

    filas_limpias = []
    for fila in tabla_principal:
        fila_txt = [limpiar_celda(c) for c in fila]
        if any(fila_txt):
            filas_limpias.append(fila_txt)

    idx_header = 0
    for i, fila in enumerate(filas_limpias):
        no_vacias = sum(1 for c in fila if c != "")
        if no_vacias >= 4:
            idx_header = i
            break

    raw_headers = filas_limpias[idx_header]
    headers = []
    for j, h in enumerate(raw_headers):
        if j == 0:
            headers.append(h if h else "Origin")
        else:
            headers.append(h if h else f"Dest_{j}")

    datos = []
    for fila in filas_limpias[idx_header + 1 :]:
        fila_ajustada = (fila + [""] * len(headers))[: len(headers)]
        origen = fila_ajustada[0]
        if not origen or origen.startswith(("*", "†", "§", "Note", "Example")):
            continue
        valores = [origen] + [parsear_valor(v) for v in fila_ajustada[1:]]
        datos.append(valores)

    df_matriz = pd.DataFrame(datos, columns=headers)

    df_plana = df_matriz.melt(
        id_vars=[headers[0]],
        var_name="Destination",
        value_name="Surcharge_USD_kg",
    ).rename(columns={headers[0]: "Origin"})
    df_plana = df_plana[df_plana["Surcharge_USD_kg"].notna() & (df_plana["Surcharge_USD_kg"] != "")]

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df_matriz.to_excel(writer, sheet_name="Matriz_Original", index=False)
        df_plana.to_excel(writer, sheet_name="Tabla_Plana", index=False)

        for sheet_name in ("Matriz_Original", "Tabla_Plana"):
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "B2" if sheet_name == "Matriz_Original" else "A2"
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)


def main():
    forzar = os.environ.get("FORZAR_ENVIO", "").strip().lower() == "true"

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

    hubo_cambio = nuevo_hash != hash_anterior

    if not hubo_cambio and not forzar:
        print("El archivo no cambió desde la última revisión. No se envía correo.")
        escribir_output("enviar", "false")
        return

    ruta_pdf = os.path.join(CARPETA_DESTINO, f"{nombre_base}{ext}")
    ruta_excel = os.path.join(CARPETA_DESTINO, f"{nombre_base}.xlsx")

    with open(ruta_pdf, "wb") as f:
        f.write(contenido)

    pdf_a_excel(ruta_pdf, ruta_excel)

    with open(ARCHIVO_HASH, "w", encoding="utf-8") as f:
        f.write(nuevo_hash)

    motivo = "Cambio detectado" if hubo_cambio else "Envío forzado manualmente"
    print(f"{motivo}. PDF: {ruta_pdf} | Excel: {ruta_excel}")
    escribir_output("enviar", "true")
    escribir_output("archivo_path", ruta_pdf)
    escribir_output("excel_path", ruta_excel)
    escribir_output("archivo_nombre", f"{nombre_base}{ext}")


if __name__ == "__main__":
    main()
