import hashlib
import mimetypes
import os
import re
import smtplib
from email.message import EmailMessage
from urllib.parse import urljoin, urlparse

import pandas as pd
import pdfplumber
from playwright.sync_api import sync_playwright

URL = "https://gffsolutions.ups.com/IAF-Surcharges.html"
CARPETA_DESTINO = "descargas"
ARCHIVO_HASH = os.path.join(CARPETA_DESTINO, "ultimo.sha256")
os.makedirs(CARPETA_DESTINO, exist_ok=True)

# Normalización de nombres de origen que el PDF suele partir por espaciado de fuente
ORIGENES_CANONICOS = {
    "northchina": "North China*",
    "eastchina": "East China**",
    "southchina": "South China***",
    "hongkong": "Hong Kong",
    "korea": "Korea",
    "taiwan": "Taiwan",
    "singapore": "Singapore",
    "thailand": "Thailand",
    "malaysia": "Malaysia",
    "vietnamhochiminh": "Vietnam - Ho Chi Minh",
    "vietnamhanoi": "Vietnam - Hanoi",
    "philippines": "Philippines",
    "cambodia": "Cambodia",
    "myanmar": "Myanmar",
    "newzealand": "New Zealand",
    "indonesia": "Indonesia",
    "australia": "Australia",
    "unitedstates": "United States",
    "canada": "Canada",
    "unitedkingdomireland": "United Kingdom, Ireland",
    "germanyitalyturkey": "Germany, Italy, Turkey",
    "restofeurope": "Rest of Europe",
    "india": "India",
    "bangladesh": "Bangladesh",
    "egypt": "Egypt",
    "southafrica": "South Africa",
    "africaexceptforsouthafrica": "Africa (Except for South Africa)",
}

# Respaldo de códigos IATA de China según notas al pie del PDF de UPS
CHINA_IATA_DEFAULT = {
    "North China*": ["PEK", "DLC", "TAO", "SHE", "CGO", "TSN", "SIA"],
    "East China**": ["CTU", "CKG", "HGH", "KUB", "NKG", "NGB", "SHA", "SZV", "WUX"],
    "South China***": ["XMN", "FOC", "CAN", "SZX"],
}

# Mapeo de orígenes de UPS a códigos IATA
ORIGEN_A_IATA = {
    "Hong Kong": ["HKG"],
    "Korea": ["ICN"],
    "Taiwan": ["TPE"],
    "Singapore": ["SIN"],
    "Thailand": ["BKK"],
    "Malaysia": ["KUL"],
    "Vietnam - Ho Chi Minh": ["SGN"],
    "Vietnam - Hanoi": ["HAN"],
    "Philippines": ["MNL"],
    "Cambodia": ["PNH"],
    "Myanmar": ["RGN"],
    "New Zealand": ["AKL"],
    "Indonesia": ["CGK"],
    "Australia": ["SYD", "MEL"],
    "United States": ["LAX", "MIA"],
    "Canada": ["YYZ", "YVR"],
    "United Kingdom, Ireland": ["LHR", "DUB"],
    "Germany, Italy, Turkey": ["FRA", "MXP"],  # Nota PDF: Excluye TR
    "Rest of Europe": ["MAD", "AMS", "CDG"],
    "India": ["DEL", "BOM"],
    "Bangladesh": ["DAC"],
    "Egypt": ["CAI"],
    "South Africa": ["JNB"],
    "Africa (Except for South Africa)": ["NBO", "LOS"],
}

# Mapeo de columnas de destino de UPS a tus aeropuertos IATA operativos
DESTINO_A_IATA = {
    "US": ["LAX", "MIA"],
    "PE": ["LIM"],
    "AR, BR, CL, PY, UY": ["ASU", "MVD", "SCL"],  # BUE tiene su propia columna dedicada
    "Buenos Aires (BUE)": ["BUE"],
    "Rest of Americas": ["BOG", "GUA", "GYE", "MGA", "PTY", "SAL", "SAP", "SRZ", "TGU", "UIO"],
    "Canada": ["YYZ", "YVR"],
    "India, Pakistan, Bangladesh": ["DEL", "BOM", "DAC"],
    "Africa, South Africa": ["JNB", "CAI"],
    "Europe": ["FRA", "MAD", "AMS", "LHR", "MXP"],
    "APAC": ["HKG", "ICN", "TPE", "SIN", "BKK", "KUL", "SGN", "HAN", "MNL", "CGK", "SYD"],
}


def limpiar_nombre(texto: str) -> str:
    limpio = re.sub(r"[^a-zA-Z0-9_-]+", "_", texto).strip("_")
    return limpio[:80] or "ups_iaf_surcharge"


def calcular_sha256(datos: bytes) -> str:
    return hashlib.sha256(datos).hexdigest()


def parsear_monto(texto: str):
    if not texto:
        return None
    limpio = texto.replace("$", "").replace(",", ".").strip()
    try:
        return float(limpio)
    except ValueError:
        return texto


def normalizar_origen(texto: str) -> str:
    clave = re.sub(r"[^a-zA-Z]", "", texto).lower()
    return ORIGENES_CANONICOS.get(clave, texto.strip())


def _agrupar_por_y(palabras: list[dict], tolerancia: float = 4.0) -> list[list[dict]]:
    if not palabras:
        return []
    ordenadas = sorted(palabras, key=lambda w: (w["top"], w["x0"]))
    filas = [[ordenadas[0]]]
    for w in ordenadas[1:]:
        if abs(w["top"] - filas[-1][0]["top"]) <= tolerancia:
            filas[-1].append(w)
        else:
            filas.append([w])
    return filas


def _extraer_iatas_china_notas(palabras_notas: list[dict]) -> dict[str, list[str]]:
    """Lee dinámicamente los códigos IATA de North/East/South China desde las notas al pie."""
    resultado = {k: list(v) for k, v in CHINA_IATA_DEFAULT.items()}
    lineas = _agrupar_por_y(palabras_notas, tolerancia=4.0)
    for linea in lineas:
        txt = " ".join(w["text"] for w in sorted(linea, key=lambda x: x["x0"]))
        codigos = re.findall(r"\b[A-Z]{3}\b", txt)
        if not codigos:
            continue
        if "North China" in txt:
            resultado["North China*"] = codigos
        elif "East China" in txt:
            resultado["East China**"] = codigos
        elif "South China" in txt:
            resultado["South China***"] = codigos
    return resultado


def _mapear_destino_iatas(col_destino: str) -> list[str]:
    if col_destino in DESTINO_A_IATA:
        return DESTINO_A_IATA[col_destino]
    col_low = col_destino.lower()
    if col_low == "us":
        return DESTINO_A_IATA["US"]
    if col_low == "pe":
        return DESTINO_A_IATA["PE"]
    if "bue" in col_low:
        return DESTINO_A_IATA["Buenos Aires (BUE)"]
    if "ar," in col_low or "py" in col_low or "uy" in col_low:
        return DESTINO_A_IATA["AR, BR, CL, PY, UY"]
    if "rest of americas" in col_low:
        return DESTINO_A_IATA["Rest of Americas"]
    if "canada" in col_low:
        return DESTINO_A_IATA["Canada"]
    if "india" in col_low:
        return DESTINO_A_IATA["India, Pakistan, Bangladesh"]
    if "africa" in col_low:
        return DESTINO_A_IATA["Africa, South Africa"]
    if "europe" in col_low:
        return DESTINO_A_IATA["Europe"]
    if "apac" in col_low:
        return DESTINO_A_IATA["APAC"]
    return [col_destino]


def pdf_a_excel(ruta_pdf: str, ruta_excel: str):
    with pdfplumber.open(ruta_pdf) as pdf:
        page = pdf.pages[0]
        palabras = page.extract_words(
            keep_blank_chars=False,
            use_text_flow=False,
            x_tolerance=5,
            y_tolerance=3,
        )

    if not palabras:
        raise RuntimeError("El PDF no contiene texto seleccionable.")

    origins_words = [w for w in palabras if w["text"] == "Origins"]
    north_words = [w for w in palabras if w["text"].startswith("North") and not w["text"].startswith("*")]
    notas_words = [
        w for w in palabras
        if w["text"].startswith(("*", "†", "§", "∞")) or w["text"] == "Example"
    ]

    if not origins_words or not north_words:
        raise RuntimeError("No se encontraron los límites 'Origins' / 'North China' en el PDF.")

    top_header = min(w["top"] for w in origins_words) - 12.0
    top_datos = min(w["top"] for w in north_words) - 2.0
    bottom_datos = min((w["top"] for w in notas_words if w["top"] > top_datos), default=page.height) - 3.0

    palabras_header = [
        w for w in palabras
        if top_header <= w["top"] < top_datos
        and w["text"] not in {"Destinations", "Surcharge", "Amount", "(USD", "per", "kilo)†"}
    ]
    palabras_datos = [
        w for w in palabras
        if top_datos <= w["top"] <= bottom_datos
    ]
    palabras_pie = [
        w for w in palabras
        if bottom_datos < w["top"] <= bottom_datos + 95.0
    ]

    china_iatas = _extraer_iatas_china_notas(palabras_pie)

    tokens_precio = [w for w in palabras_datos if w["text"].startswith("$")]
    if not tokens_precio:
        raise RuntimeError("No se encontraron montos en dólares dentro de la tabla.")

    xs_precio = sorted((w["x0"] + w["x1"]) / 2.0 for w in tokens_precio)
    clusters_x: list[list[float]] = [[xs_precio[0]]]
    for x in xs_precio[1:]:
        if abs(x - (sum(clusters_x[-1]) / len(clusters_x[-1]))) <= 14.0:
            clusters_x[-1].append(x)
        else:
            clusters_x.append([x])

    centros_dest = [sum(c) / len(c) for c in clusters_x]

    limite_izq_dest = centros_dest[0] - 20.0
    limite_der_dest = centros_dest[-1] + 20.0
    cortes_dest = [
        (centros_dest[i] + centros_dest[i + 1]) / 2.0
        for i in range(len(centros_dest) - 1)
    ]

    def indice_columna(w: dict) -> int:
        xc = (w["x0"] + w["x1"]) / 2.0
        if xc < limite_izq_dest:
            return 0
        if xc > limite_der_dest:
            return len(centros_dest) + 1
        for idx, corte in enumerate(cortes_dest):
            if xc < corte:
                return idx + 1
        return len(centros_dest)

    total_cols = len(centros_dest) + 2

    buckets_header: list[list[dict]] = [[] for _ in range(total_cols)]
    for w in palabras_header:
        buckets_header[indice_columna(w)].append(w)

    headers = []
    for idx, bucket in enumerate(buckets_header):
        if idx == 0:
            headers.append("Origins")
        elif idx == total_cols - 1:
            headers.append("Effective Date")
        else:
            bucket_ord = sorted(bucket, key=lambda w: (round(w["top"] / 4.0), w["x0"]))
            nombre = " ".join(w["text"] for w in bucket_ord).strip()
            headers.append(nombre if nombre else f"Dest_{idx}")

    filas_agrupadas = _agrupar_por_y(palabras_datos, tolerancia=4.0)
    registros = []
    for fila_words in filas_agrupadas:
        celdas: list[list[dict]] = [[] for _ in range(total_cols)]
        for w in fila_words:
            celdas[indice_columna(w)].append(w)

        origen_raw = " ".join(w["text"] for w in sorted(celdas[0], key=lambda x: x["x0"])).strip()
        if not origen_raw:
            continue
        origen = normalizar_origen(origen_raw)

        fila_final = [origen]
        for col_idx in range(1, total_cols - 1):
            txt_val = " ".join(w["text"] for w in sorted(celdas[col_idx], key=lambda x: x["x0"])).strip()
            fila_final.append(parsear_monto(txt_val) if txt_val else None)

        eff_date = " ".join(w["text"] for w in sorted(celdas[-1], key=lambda x: x["x0"])).strip()
        fila_final.append(eff_date)
        registros.append(fila_final)

    df_matriz = pd.DataFrame(registros, columns=headers)

    cols_destinos = headers[1:-1]
    df_plana = df_matriz.melt(
        id_vars=["Origins", "Effective Date"],
        value_vars=cols_destinos,
        var_name="Destination",
        value_name="Surcharge_USD_kg",
    )
    df_plana = df_plana[df_plana["Surcharge_USD_kg"].notna()][
        ["Origins", "Destination", "Surcharge_USD_kg", "Effective Date"]
    ].reset_index(drop=True)

    # Construcción de Tabla_IATA expandiendo China (notas al pie) y tus destinos IATA
    mapa_origen_completo = {**china_iatas, **ORIGEN_A_IATA}
    filas_iata = []
    for _, row in df_plana.iterrows():
        orig_ups = row["Origins"]
        dest_ups = row["Destination"]
        monto = row["Surcharge_USD_kg"]
        eff_date = row["Effective Date"]

        orig_iatas = mapa_origen_completo.get(orig_ups, [orig_ups])
        dest_iatas = _mapear_destino_iatas(dest_ups)

        for o_iata in orig_iatas:
            for d_iata in dest_iatas:
                filas_iata.append(
                    {
                        "Origin_IATA": o_iata,
                        "Destination_IATA": d_iata,
                        "Surcharge_USD_kg": monto,
                        "Effective_Date": eff_date,
                        "Origin_UPS": orig_ups,
                        "Destination_UPS": dest_ups,
                    }
                )

    df_iata = pd.DataFrame(filas_iata).drop_duplicates(
        subset=["Origin_IATA", "Destination_IATA", "Origin_UPS", "Destination_UPS"]
    )

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df_iata.to_excel(writer, sheet_name="Tabla_IATA", index=False)
        df_plana.to_excel(writer, sheet_name="Tabla_Plana", index=False)
        df_matriz.to_excel(writer, sheet_name="Matriz_Original", index=False)

        for sheet_name in ("Tabla_IATA", "Tabla_Plana", "Matriz_Original"):
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "B2" if sheet_name == "Matriz_Original" else "A2"
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 32)


def enviar_correo(ruta_pdf: str, ruta_excel: str, nombre_base: str):
    mail_user = os.environ.get("MAIL_USER", "").strip()
    mail_pass = os.environ.get("MAIL_PASS", "").replace(" ", "").strip()
    destinatario = os.environ.get("DESTINATARIO", "Jose.Mon@mirgor.com").strip()

    if not mail_user or not mail_pass:
        raise RuntimeError("Faltan los secrets MAIL_USER o MAIL_PASS en GitHub Actions.")

    msg = EmailMessage()
    msg["Subject"] = f"Actualización UPS IAF Surcharge - {nombre_base}"
    msg["From"] = mail_user
    msg["To"] = destinatario
    msg.set_content(
        "Se adjuntan el PDF original de UPS y el Excel procesado con códigos IATA "
        "(hojas Tabla_IATA, Tabla_Plana y Matriz_Original)."
    )

    for ruta in (ruta_pdf, ruta_excel):
        ctype, _ = mimetypes.guess_type(ruta)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        with open(ruta, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(ruta),
            )

    print(f"Enviando correo desde {mail_user} hacia {destinatario}...")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(mail_user, mail_pass)
        smtp.send_message(msg)
    print("Correo enviado correctamente.")


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
            with open(download.path(), "rb") as f:
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
        return

    ruta_pdf = os.path.join(CARPETA_DESTINO, f"{nombre_base}{ext}")
    ruta_excel = os.path.join(CARPETA_DESTINO, f"{nombre_base}.xlsx")

    with open(ruta_pdf, "wb") as f:
        f.write(contenido)

    pdf_a_excel(ruta_pdf, ruta_excel)
    enviar_correo(ruta_pdf, ruta_excel, nombre_base)

    with open(ARCHIVO_HASH, "w", encoding="utf-8") as f:
        f.write(nuevo_hash)


if __name__ == "__main__":
    main()
