def pdf_a_excel(ruta_pdf: str, ruta_excel: str):
    filas_tabla = []

    with pdfplumber.open(ruta_pdf) as pdf:
        page = pdf.pages[0]
        palabras = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        if not palabras:
            raise RuntimeError("El PDF no contiene texto extraíble.")

        # 1. Ubicamos la fila de encabezados de destinos ("US", "PE", ..., "Effective Date")
        # y la fila de "Origins" para fijar el límite superior y las coordenadas X de cada columna.
        candidatos_header = [
            w for w in palabras
            if w["text"] in {"US", "PE", "Canada", "Europe", "APAC", "Effective"}
        ]
        if not candidatos_header:
            raise RuntimeError("No se encontraron encabezados conocidos de la tabla en el PDF.")

        top_header_min = min(w["top"] for w in candidatos_header) - 6.0
        top_header_max = max(w["bottom"] for w in candidatos_header) + 6.0

        # Límite inferior: justo antes de las notas al pie ("* North China" o "Example Shipment")
        notas = [
            w for w in palabras
            if w["text"].startswith(("*", "†", "§", "∞")) or w["text"] == "Example"
        ]
        bottom_limit = min((w["top"] for w in notas), default=page.height) - 4.0

        # Recortamos la zona de la tabla para ignorar títulos superiores y notas inferiores
        area_tabla = page.crop((0, top_header_min, page.width, bottom_limit))
        tablas = area_tabla.extract_tables(
            table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 4,
                "join_tolerance": 4,
            }
        )

        if tablas:
            tabla_bruta = max(tablas, key=lambda t: len(t) * (len(t[0]) if t else 0))
            for fila in tabla_bruta:
                fila_limpia = [limpiar_celda(c) for c in fila]
                if any(fila_limpia):
                    filas_tabla.append(fila_limpia)
        else:
            # Respaldo por coordenadas dentro del recorte de la tabla
            palabras_recorte = area_tabla.extract_words(keep_blank_chars=False, use_text_flow=False)
            palabras_ordenadas = sorted(palabras_recorte, key=lambda w: (w["top"], w["x0"]))
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
                    if (w["x0"] - prev_x1) > 10.0:
                        celdas.append(" ".join(actual))
                        actual = [w["text"]]
                    else:
                        actual.append(w["text"])
                    prev_x1 = max(prev_x1, w["x1"])
                celdas.append(" ".join(actual))
                if len(celdas) >= 2:
                    filas_tabla.append(celdas)

    if len(filas_tabla) < 2:
        raise RuntimeError("No se pudo reconstruir la tabla del recuadro.")

    # Buscamos la fila que contiene "Origins" o "US" como fila de encabezados
    idx_header = 0
    for i, fila in enumerate(filas_tabla):
        texto_fila = " ".join(fila).lower()
        if "origins" in texto_fila or ("us" in [c.lower() for c in fila] and "apac" in texto_fila):
            idx_header = i
            break

    raw_headers = filas_tabla[idx_header]
    headers = []
    for j, h in enumerate(raw_headers):
        if j == 0:
            headers.append("Origins")
        else:
            headers.append(h if h else f"Col_{j}")

    # Evitamos nombres de columna duplicados si los hubiera
    headers_unicos = []
    vistos = {}
    for h in headers:
        if h in vistos:
            vistos[h] += 1
            headers_unicos.append(f"{h}_{vistos[h]}")
        else:
            vistos[h] = 1
            headers_unicos.append(h)

    filas_datos = []
    for fila in filas_tabla[idx_header + 1 :]:
        fila_ajustada = (fila + [""] * len(headers_unicos))[: len(headers_unicos)]
        origen = fila_ajustada[0].strip()
        if not origen or origen.lower() in {"origins", "destinations", "surcharge amount (usd per kilo)†"}:
            continue
        if origen.startswith(("*", "†", "§", "∞", "Example")):
            break

        # En los valores de surcharge, si una celda trae dos montos pegados (ej. "$2.66 $2.66"), tomamos el primero
        valores = [origen]
        for idx_col, val in enumerate(fila_ajustada[1:], start=1):
            col_nombre = headers_unicos[idx_col].lower()
            if "effective" in col_nombre or "date" in col_nombre:
                valores.append(val)
            else:
                primer_token = val.split()[0] if val else ""
                valores.append(parsear_valor(primer_token))
        filas_datos.append(valores)

    df_matriz = pd.DataFrame(filas_datos, columns=headers_unicos)

    # Excluimos Effective Date del unpivot de tarifas para que Surcharge_USD_kg sea numérico limpio
    cols_tarifa = [
        c for c in headers_unicos[1:]
        if "effective" not in c.lower() and "date" not in c.lower()
    ]

    df_plana = df_matriz.melt(
        id_vars=[headers_unicos[0]],
        value_vars=cols_tarifa,
        var_name="Destination",
        value_name="Surcharge_USD_kg",
    ).rename(columns={headers_unicos[0]: "Origins"})

    df_plana = df_plana[
        df_plana["Surcharge_USD_kg"].notna() & (df_plana["Surcharge_USD_kg"] != "")
    ]

    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
        df_matriz.to_excel(writer, sheet_name="Matriz_Original", index=False)
        df_plana.to_excel(writer, sheet_name="Tabla_Plana", index=False)

        for sheet_name in ("Matriz_Original", "Tabla_Plana"):
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "B2" if sheet_name == "Matriz_Original" else "A2"
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 35)
