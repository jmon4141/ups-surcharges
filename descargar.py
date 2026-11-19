def _extraer_por_coordenadas(pdf) -> list[list[str]]:
    """Reconstruye filas y columnas aproximadas usando coordenadas (x0, top) de las palabras."""
    filas_reconstruidas = []
    for page in pdf.pages:
        palabras = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        if not palabras:
            continue

        # Agrupamos palabras que están en la misma línea horizontal (tolerancia de 4 pt en 'top')
        palabras_ordenadas = sorted(palabras, key=lambda w: (w["top"], w["x0"]))
        lineas = []
        for w in palabras_ordenadas:
            if not lineas or abs(w["top"] - lineas[-1][-1]["top"]) > 4.0:
                lineas.append([w])
            else:
                lineas[-1].append(w)

        for linea in lineas:
            linea_sorted = sorted(linea, key=lambda w: w["x0"])
            # Agrupamos palabras cercanas en X dentro de una misma celda (separación > 14 pt = nueva columna)
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
        # Fallback: guardamos al menos el texto crudo para no abortar la notificación
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
