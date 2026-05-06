from __future__ import annotations
import os
import re
import shutil
import tempfile
import numpy as np
import pandas as pd
from datetime import date, datetime, timedelta
from pdf2image import convert_from_path, convert_from_bytes
import pytesseract
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except Exception:
    _HAS_PDFPLUMBER = False
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# Detectar dónde está poppler (pdftoppm) en este sistema
def _find_poppler():
    for cmd in ("pdftoppm", "pdfinfo", "pdftocairo"):
        p = shutil.which(cmd)
        if p:
            return os.path.dirname(p)
    for candidate in ("/usr/bin", "/usr/local/bin", "/opt/homebrew/bin"):
        if os.path.exists(os.path.join(candidate, "pdftoppm")):
            return candidate
    return None

_POPPLER_PATH = _find_poppler()

st.set_page_config(
    page_title="Conciliación Bancaria",
    page_icon="🏦",
    layout="centered",
)

st.title("🏦 Conciliación Bancaria")

col_a, col_b = st.columns(2)
uploaded_ext = col_a.file_uploader("📄 Extracto bancario (PDF)", type="pdf")
uploaded_aux = col_b.file_uploader("📒 Auxiliar contabilidad (PDF)", type="pdf")

if not uploaded_ext or not uploaded_aux:
    st.info("Sube los dos archivos PDF para continuar.")
    st.stop()

col_a.success(uploaded_ext.name)
col_b.success(uploaded_aux.name)

if st.button("Generar conciliación", type="primary"):

    with st.spinner("Procesando... (puede tardar 1-2 minutos)"):

        # ── Guardar PDFs temporales ───────────────────────────────────────────
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_ext.read())
            pdf_ext_path = tmp.name
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_aux.read())
            pdf_aux_path = tmp.name

        # ── Utilidades ────────────────────────────────────────────────────────

        def _parse_num(text):
            clean = re.sub(r"[^\d.]", "", str(text))
            try:    return float(clean) if clean else None
            except: return None

        def _is_currency(text):
            t = re.sub(r"[\[\|\]]+", "", text).strip()
            return bool(
                re.match(r"^-?\d{1,3}(,\d{3})+(\.\d{2})?$", t) or
                re.match(r"^-?\d{4,}(\.\d{2})?$", t)
            )

        def _fmt_cop(v):
            if v is None or (isinstance(v, float) and np.isnan(v)): return ""
            try:    return f"$ {float(v):>20,.0f}"
            except: return str(v)

        def _plumber_lines(page):
            """Extrae líneas desde pdfplumber (sin OCR). Retorna (lines, page_width)."""
            words = page.extract_words(x_tolerance=3, y_tolerance=3,
                                       keep_blank_chars=False)
            if not words:
                return [], page.width
            lines, cur, cur_y = [], [], None
            for w in sorted(words, key=lambda w: (round(w["top"], 1), w["x0"])):
                if cur_y is None or abs(w["top"] - cur_y) <= 4:
                    cur.append((w["x0"], w["text"]))
                    if cur_y is None: cur_y = w["top"]
                else:
                    if cur: lines.append(sorted(cur))
                    cur, cur_y = [(w["x0"], w["text"])], w["top"]
            if cur: lines.append(sorted(cur))
            return lines, page.width

        def _ocr_lines(img):
            df = pytesseract.image_to_data(
                img, lang="spa", config="--psm 6",
                output_type=pytesseract.Output.DATAFRAME
            )
            df = df[df.conf > 0].dropna(subset=["text"]).copy()
            df["text"] = df["text"].astype(str).str.strip()
            df = df[df["text"] != ""].sort_values(["top", "left"])
            lines, cur, cur_y = [], [], None
            for _, row in df.iterrows():
                if cur_y is None or abs(row["top"] - cur_y) <= 12:
                    cur.append((row["left"], row["text"]))
                    cur_y = cur_y or row["top"]
                else:
                    if cur: lines.append(sorted(cur))
                    cur, cur_y = [(row["left"], row["text"])], row["top"]
            if cur: lines.append(sorted(cur))
            return lines

        # ── Extracto: parser genérico (Itau, Banco de Bogotá, etc.) ─────────────
        def _clean(t: str) -> str:
            """Elimina artefactos OCR como | al inicio/fin."""
            return re.sub(r"[\[\|\]]+", "", t).strip()

        def _is_amount(text: str) -> bool:
            t = _clean(text)
            return bool(
                # Con agrupación de comas → decimal opcional (OCR a veces lo pierde)
                re.match(r"^-?\d{1,3}(,\d{3})+(\.\d{1,2})?$", t) or
                # Sin comas → exige decimal para no confundir con referencias
                re.match(r"^-?\d{4,}\.\d{2}$", t)
            )

        def _to_float(text: str):
            t = _clean(text).replace(",", "")
            try:    return float(t)
            except: return None

        def _parse_extracto_page(lines, w):
            rows = []
            for line in lines:
                if not line: continue

                # Detectar fecha en los primeros tokens (dentro del 15% izquierdo)
                # Acepta: DD/MM (Banco de Bogotá), DD solo (Itau)
                day = None
                mon = None
                for tok_x, tok_t in line:
                    if tok_x >= w * 0.15:
                        break          # ya salimos de la zona de fecha
                    tc = _clean(tok_t)
                    # DD/MM
                    m = re.match(r"^(\d{1,2})[/\-](\d{2})$", tc)
                    if m:
                        d = int(m.group(1)); mo = int(m.group(2))
                        if 1 <= d <= 31 and 1 <= mo <= 12:
                            day = d; mon = mo; break
                    # Solo DD
                    if re.match(r"^\d{1,2}$", tc):
                        d = int(tc)
                        if 1 <= d <= 31:
                            day = d; break

                if day is None: continue

                # Dos zonas por posición X:
                #   Transacción (cargo/abono): 55 %–80 % del ancho
                #   Saldo:                     > 80 % del ancho
                # Esto evita confundir NITs, referencias y saldos entre sí.
                txn_amounts   = []   # (x, valor) en zona cargo/abono
                saldo_amounts = []   # (x, valor) en zona saldo
                i = 0
                tokens = line
                while i < len(tokens):
                    x, t = tokens[i]
                    tc = _clean(t)
                    if x < w * 0.55:          # zona de descripción → ignorar
                        i += 1; continue
                    # Monto negativo partido: "-" y "1,234.00"
                    if tc == "-" and i + 1 < len(tokens):
                        nx, nt = tokens[i+1]
                        if _is_amount(nt) and nx >= w * 0.55 and (nx - x) < w * 0.05:
                            v = _to_float(nt)
                            if v:
                                if nx > w * 0.80: saldo_amounts.append((nx, -abs(v)))
                                else:             txn_amounts.append((nx, -abs(v)))
                            i += 2; continue
                    if _is_amount(tc):
                        v = _to_float(tc)
                        if v is not None and abs(v) >= 100:   # excluir artefactos menores a $100
                            if x > w * 0.80: saldo_amounts.append((x, v))
                            else:             txn_amounts.append((x, v))
                    i += 1

                if not saldo_amounts or not txn_amounts: continue

                # Saldo: monto más a la derecha en la zona de saldo
                saldo_amounts.sort(key=lambda a: a[0])
                saldo = abs(saldo_amounts[-1][1])

                # Transacción: monto más a la derecha en la zona cargo/abono
                txn_amounts.sort(key=lambda a: a[0])
                x_valor = txn_amounts[-1][0]
                valor   = txn_amounts[-1][1]

                # Clasificar cargo vs abono:
                #   Banco de Bogotá (columna única "Valor"): negativo = cargo, positivo = abono
                #   Itaú (columnas separadas): posición determina la dirección
                if valor < 0:
                    # Signo negativo → siempre cargo (salida)
                    debito, credito = abs(valor), 0.0
                else:
                    # Valor positivo: usar palabras clave de la descripción
                    desc_tokens = [t for x, t in line if w*0.10 < x < w*0.65 and not _is_amount(t)]
                    desc_lower  = " ".join(desc_tokens).lower()
                    # Palabras que indican CARGO (salida de dinero)
                    cargo_kw = {"cargo", "giro", "gravamen", " db ", "pago",
                                "impuesto", "cheque", "comision"}
                    if any(k in desc_lower for k in cargo_kw):
                        debito, credito = valor, 0.0
                    else:
                        # Por defecto positivo = abono (entrada de dinero)
                        # Cubre: "Cr Ach", "Abono", "consignacion", transferencias entrantes
                        debito, credito = 0.0, valor

                # Descripción: tokens entre fecha y el inicio de la zona de transacción
                min_txn_x = txn_amounts[0][0]
                desc_parts = [
                    _clean(t) for x, t in line
                    if w * 0.10 < x < min_txn_x - w * 0.02
                    and not _is_amount(t)
                    and len(_clean(t)) > 1
                ]
                desc = " ".join(desc_parts[:10])

                rows.append({
                    "dia": day, "mes": mon, "descripcion": desc,
                    "debito": debito, "credito": credito, "saldo": saldo,
                })
            return rows

        def _saldo_final_extracto(images) -> float | None:
            """Busca 'Saldo Final' en el texto del extracto, desde la última página."""
            for img in reversed(images):
                text = pytesseract.image_to_string(img, lang="spa", config="--psm 6")
                for line in text.split("\n"):
                    if re.search(r"saldo\s+(final|al\s+corte|bancario)", line, re.IGNORECASE):
                        nums = re.findall(r"[\d,]+\.\d{2}", line)
                        for n in nums:
                            v = _to_float(n)
                            if v and v > 0:
                                return v
            return None

        def _pdf_to_images(pdf_path, dpi=150):
            """Convierte PDF a imágenes PIL sin necesitar poppler del sistema."""
            # Intento 1: pdfplumber (usa pypdfium2 internamente en v0.11+)
            try:
                imgs = []
                with pdfplumber.open(pdf_path) as pdf:
                    for page in pdf.pages:
                        imgs.append(page.to_image(resolution=dpi).original)
                if imgs:
                    return imgs
            except Exception:
                pass
            # Intento 2: pdf2image con poppler
            try:
                with open(pdf_path, "rb") as f:
                    return convert_from_bytes(f.read(), dpi=dpi,
                                              poppler_path=_POPPLER_PATH)
            except Exception:
                pass
            return []

        def parse_extracto(ano, mes):
            # pdfplumber: lee texto directo del PDF (sin OCR, sin poppler)
            all_rows = []
            filas_por_pagina = []
            images = None
            plumber_ok = False
            plumber_sample = []   # primeras líneas de página 1 para debug
            try:
                with pdfplumber.open(pdf_ext_path) as pdf:
                    for pg, page in enumerate(pdf.pages, 1):
                        lines, w = _plumber_lines(page)
                        if pg == 1:
                            plumber_sample = [
                                " | ".join(t for _, t in ln)
                                for ln in lines[:8]
                            ]
                        page_rows = _parse_extracto_page(lines, w)
                        filas_por_pagina.append((pg, len(page_rows), "pdf"))
                        all_rows.extend(page_rows)
                plumber_ok = True   # terminó sin excepción
            except Exception as e:
                all_rows = []
                plumber_sample = [f"ERROR: {e}"]

            st.session_state["plumber_sample"] = plumber_sample

            # Solo usar OCR si pdfplumber lanzó excepción (no por pocas filas)
            if not plumber_ok:
                images = _pdf_to_images(pdf_ext_path, dpi=150)
                for pg, img in enumerate(images, 1):
                    try:
                        page_rows = _parse_extracto_page(_ocr_lines(img), img.width)
                    except Exception:
                        page_rows = []
                    filas_por_pagina.append((pg, len(page_rows), "ocr"))
                    all_rows.extend(page_rows)
                st.session_state["ext_debug"] = filas_por_pagina

            if not all_rows:
                return pd.DataFrame(), images
            df = pd.DataFrame(all_rows)
            def _make_date(row):
                m = int(row["mes"]) if pd.notna(row.get("mes")) and row.get("mes") else mes
                max_day = 28 if m == 2 else 30 if m in [4,6,9,11] else 31
                return date(ano, m, min(int(row["dia"]), max_day))
            df["fecha"] = df.apply(_make_date, axis=1)
            return df.sort_values("fecha").reset_index(drop=True), images

        # ── Informe Movimiento General (última página) ────────────────────────
        # Sin posiciones fijas: usa delta de saldo para determinar débito/crédito
        def parse_informe(img):
            lines = _ocr_lines(img)
            rows = []
            prev_saldo = None
            for line in lines:
                # Línea de asiento: contiene patrón L-1-919-1 etc.
                comprobante = next((t for _, t in line if re.match(r"^[A-Z]-\d+-\d+", t)), None)
                if not comprobante: continue
                fecha_str = next((t for _, t in line if re.match(r"^\d{4}/\d{2}/\d{2}$", t)), None)
                if not fecha_str: continue
                try:    fecha = datetime.strptime(fecha_str, "%Y/%m/%d").date()
                except: continue

                # Montos solo en zona derecha (x > 50%): evita capturar números de descripción
                amounts = sorted(
                    [(x, _parse_num(t)) for x, t in line
                     if _is_currency(t) and _parse_num(t) and _parse_num(t) > 0
                     and x > img.width * 0.50],
                    key=lambda a: a[0]
                )
                if not amounts: continue

                # El saldo es siempre el monto más a la derecha
                saldo = amounts[-1][1]

                # El monto de transacción es el anterior al saldo
                non_saldo = [v for _, v in amounts[:-1]]
                if not non_saldo: continue
                amount = non_saldo[-1]

                # Débito = saldo sube (entra dinero), Crédito = saldo baja (sale dinero)
                if prev_saldo is not None:
                    if saldo >= prev_saldo:
                        debito, credito = amount, 0.0
                    else:
                        debito, credito = 0.0, amount
                else:
                    # Primera fila: usar posición relativa (déb está antes de 68% del ancho)
                    x_amount = amounts[-2][0] if len(amounts) >= 2 else 0
                    if x_amount < img.width * 0.68:
                        debito, credito = amount, 0.0
                    else:
                        debito, credito = 0.0, amount
                prev_saldo = saldo

                # Descripción: tokens entre la fecha y la zona de montos
                fecha_x  = next((x for x, t in line if t == fecha_str), 0)
                monto_x  = amounts[-2][0] if len(amounts) >= 2 else img.width
                desc = " ".join(
                    t for x, t in sorted(line)
                    if x > fecha_x and x < monto_x
                    and not re.match(r"^[\d,\.]+$", t)
                    and t not in ("O", "0", "SUC", "NIT")
                ).strip()

                rows.append({
                    "comprobante": comprobante, "fecha": fecha, "descripcion": desc,
                    "debito": debito, "credito": credito, "saldo": saldo,
                })
            return pd.DataFrame(rows) if rows else pd.DataFrame(
                columns=["comprobante","fecha","descripcion","debito","credito","saldo"])

        def _parse_informe_plumber(page):
            """Auxiliar con pdfplumber: misma lógica de parse_informe pero sin OCR."""
            lines, _ = _plumber_lines(page)
            rows = []
            prev_saldo = None
            for line in lines:
                comprobante = next((t for _, t in line if re.match(r"^[A-Z]-\d+-\d+", t)), None)
                if not comprobante: continue
                fecha_str = next((t for _, t in line if re.match(r"^\d{4}/\d{2}/\d{2}$", t)), None)
                if not fecha_str: continue
                try:    fecha = datetime.strptime(fecha_str, "%Y/%m/%d").date()
                except: continue
                amounts = sorted(
                    [(x, _parse_num(t)) for x, t in line
                     if _is_currency(t) and _parse_num(t) and _parse_num(t) > 0
                     and x > page.width * 0.50],
                    key=lambda a: a[0]
                )
                if not amounts: continue
                saldo = amounts[-1][1]
                non_saldo = [v for _, v in amounts[:-1]]
                if not non_saldo: continue
                amount = non_saldo[-1]
                if prev_saldo is not None:
                    if saldo >= prev_saldo: debito, credito = amount, 0.0
                    else:                   debito, credito = 0.0, amount
                else:
                    x_amt = amounts[-2][0] if len(amounts) >= 2 else 0
                    if x_amt < page.width * 0.68: debito, credito = amount, 0.0
                    else:                         debito, credito = 0.0, amount
                prev_saldo = saldo
                fecha_x = next((x for x, t in line if t == fecha_str), 0)
                monto_x  = amounts[-2][0] if len(amounts) >= 2 else page.width
                desc = " ".join(
                    t for x, t in sorted(line)
                    if x > fecha_x and x < monto_x
                    and not re.match(r"^[\d,\.]+$", t)
                    and t not in ("O", "0", "SUC", "NIT")
                ).strip()
                rows.append({"comprobante": comprobante, "fecha": fecha, "descripcion": desc,
                             "debito": debito, "credito": credito, "saldo": saldo})
            return pd.DataFrame(rows) if rows else pd.DataFrame(
                columns=["comprobante","fecha","descripcion","debito","credito","saldo"])

        def parse_auxiliar():
            """Procesa auxiliar con pdfplumber (sin OCR)."""
            all_rows = []
            aux_sample = []
            try:
                with pdfplumber.open(pdf_aux_path) as pdf:
                    for pg, page in enumerate(pdf.pages, 1):
                        if pg == 1:
                            lines_p, _ = _plumber_lines(page)
                            aux_sample = [
                                " | ".join(t for _, t in ln)
                                for ln in lines_p[:8]
                            ]
                        df_p = _parse_informe_plumber(page)
                        if not df_p.empty:
                            all_rows.append(df_p)
            except Exception as e:
                aux_sample = [f"ERROR: {e}"]

            st.session_state["aux_sample"] = aux_sample
            if not all_rows:
                return pd.DataFrame(
                    columns=["comprobante","fecha","descripcion","debito","credito","saldo"]), None
            return pd.concat(all_rows, ignore_index=True), None

        # ── Matching ──────────────────────────────────────────────────────────
        def reconciliar(df_ext, df_inf, diferencia_saldos):
            # Índice: (fecha, monto_redondeado, tipo D/C) → lista de índices
            ext_idx = {}
            for i, row in df_ext.iterrows():
                for monto, tipo in [(row.debito,"D"),(row.credito,"C")]:
                    if monto > 0:
                        ext_idx.setdefault((row.fecha, round(monto), tipo), []).append(i)

            matched_ext, matched_inf = set(), set()
            for i, row in df_inf.iterrows():
                # Débito en libros ↔ Crédito (abono) en banco
                # Crédito en libros ↔ Débito (cargo) en banco
                for monto, tipo_ext in [(row.debito,"C"),(row.credito,"D")]:
                    if monto <= 0: continue
                    found = False
                    # Tolerancia ±5 días: el banco registra en fecha valor, contabilidad
                    # registra en fecha de autorización (diferencia típica 1-5 días hábiles)
                    for day_d in [0, -1, -2, -3, -4, -5, 1, 2, 3]:
                        fecha_try = row.fecha + timedelta(days=day_d)
                        for amt_d in [0, 1, -1]:
                            key = (fecha_try, round(monto + amt_d), tipo_ext)
                            if key in ext_idx and ext_idx[key]:
                                matched_ext.add(ext_idx[key].pop(0))
                                matched_inf.add(i)
                                found = True; break
                        if found: break

            # Partidas no cruzadas del extracto → Notas no contabilizadas
            notas_cargo, notas_abono = [], []
            for i, row in df_ext.iterrows():
                if i in matched_ext: continue
                if row.debito > 0:
                    notas_cargo.append({"fecha": row.fecha, "descripcion": row.descripcion, "monto": row.debito})
                if row.credito > 0:
                    notas_abono.append({"fecha": row.fecha, "descripcion": row.descripcion, "monto": row.credito})

            # Partidas no cruzadas del auxiliar
            cheques_pend, ingresos_pend = [], []
            for i, row in df_inf.iterrows():
                if i in matched_inf: continue
                if row.credito > 0:
                    cheques_pend.append({
                        "fecha": row.fecha, "beneficiario": row.descripcion,
                        "comprobante": row.comprobante, "monto": row.credito,
                    })
                if row.debito > 0:
                    ingresos_pend.append({
                        "fecha": row.fecha, "descripcion": row.descripcion,
                        "comprobante": row.comprobante, "monto": row.debito,
                    })

            _e = lambda lst, cols: pd.DataFrame(lst) if lst else pd.DataFrame(columns=cols)
            emp = lambda cols: pd.DataFrame(columns=cols)

            # El banco agrupa pagos en lotes → el cruce 1:1 falla casi siempre.
            # Estrategia: ignorar el cruce y buscar directamente la(s) partida(s)
            # cuyo monto explica la diferencia entre saldos.
            target = round(abs(diferencia_saldos))

            if target == 0:
                # Saldos iguales → no hay partidas pendientes
                return (emp(["fecha","beneficiario","comprobante","monto"]),
                        emp(["fecha","descripcion","monto"]),
                        emp(["fecha","descripcion","monto"]),
                        emp(["fecha","descripcion","comprobante","monto"]),
                        True)

            tol = max(round(target * 0.005), 500)  # ±0.5% del target, mínimo $500

            # Candidatos: partidas cuyo monto individual ≈ target
            # Según el signo de la diferencia buscamos en las categorías correctas:
            #   diferencia < 0 (extracto < libros): cargo banco no en libros,
            #                                       o débito libros no en banco
            #   diferencia > 0 (extracto > libros): abono banco no en libros,
            #                                       o crédito libros no en banco
            if diferencia_saldos <= 0:
                candidatos = (
                    [("nc", x) for x in notas_cargo   if abs(x["monto"] - target) <= tol] +
                    [("ip", x) for x in ingresos_pend if abs(x["monto"] - target) <= tol]
                )
            else:
                candidatos = (
                    [("na", x) for x in notas_abono   if abs(x["monto"] - target) <= tol] +
                    [("ch", x) for x in cheques_pend  if abs(x["monto"] - target) <= tol]
                )

            if candidatos:
                # Elegir el más cercano al target
                cat, item = min(candidatos, key=lambda t: abs(t[1]["monto"] - target))
                return (
                    _e([item] if cat == "ch" else [], ["fecha","beneficiario","comprobante","monto"]),
                    _e([item] if cat == "nc" else [], ["fecha","descripcion","monto"]),
                    _e([item] if cat == "na" else [], ["fecha","descripcion","monto"]),
                    _e([item] if cat == "ip" else [], ["fecha","descripcion","comprobante","monto"]),
                    True,
                )

            # Si no se encontró una sola partida, intentar con ambos lados
            # (por si el signo del OCR es incorrecto)
            todos = (
                [("nc", x) for x in notas_cargo   if abs(x["monto"] - target) <= tol] +
                [("na", x) for x in notas_abono   if abs(x["monto"] - target) <= tol] +
                [("ch", x) for x in cheques_pend  if abs(x["monto"] - target) <= tol] +
                [("ip", x) for x in ingresos_pend if abs(x["monto"] - target) <= tol]
            )
            if todos:
                cat, item = min(todos, key=lambda t: abs(t[1]["monto"] - target))
                return (
                    _e([item] if cat == "ch" else [], ["fecha","beneficiario","comprobante","monto"]),
                    _e([item] if cat == "nc" else [], ["fecha","descripcion","monto"]),
                    _e([item] if cat == "na" else [], ["fecha","descripcion","monto"]),
                    _e([item] if cat == "ip" else [], ["fecha","descripcion","comprobante","monto"]),
                    True,
                )

            # No se encontró la partida explicativa
            return (
                emp(["fecha","beneficiario","comprobante","monto"]),
                emp(["fecha","descripcion","monto"]),
                emp(["fecha","descripcion","monto"]),
                emp(["fecha","descripcion","comprobante","monto"]),
                False,
            )

        # ── Excel output ──────────────────────────────────────────────────────
        _HDR  = PatternFill("solid", fgColor="1F3864")
        _YELL = PatternFill("solid", fgColor="FFF2CC")
        _RED  = PatternFill("solid", fgColor="FCE4D6")
        _GRN  = PatternFill("solid", fgColor="E2EFDA")
        _GREY = PatternFill("solid", fgColor="F2F2F2")
        _LGRY = PatternFill("solid", fgColor="D9D9D9")
        _THIN = Side(style="thin", color="CCCCCC")
        _BRD  = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
        NCOLS = 5

        def _merge_row(ws, r, text, fill, font_kw=None, height=None):
            ws.merge_cells(f"A{r}:{get_column_letter(NCOLS)}{r}")
            c = ws[f"A{r}"]
            c.value = text
            c.fill = fill
            c.alignment = Alignment(horizontal="left", vertical="center")
            kw = {"bold": True, "size": 11, "color": "1F3864"}
            if font_kw: kw.update(font_kw)
            c.font = Font(**kw)
            if height: ws.row_dimensions[r].height = height

        def _kv(ws, r, label, valor, fill_v=None):
            """Fila de clave-valor: col A = etiqueta, col B = valor."""
            ca = ws.cell(row=r, column=1, value=label)
            ca.font = Font(bold=True, size=10)
            ca.border = _BRD
            cv = ws.cell(row=r, column=2, value=valor)
            cv.border = _BRD
            cv.alignment = Alignment(horizontal="right", vertical="center")
            cv.font = Font(size=10)
            if fill_v: cv.fill = fill_v

        def _hrow(ws, r, cols):
            for c, v in enumerate(cols, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = Font(bold=True, color="FFFFFF", size=9)
                cell.fill = _HDR
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = _BRD

        def _drow(ws, r, vals, fill=None):
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                if fill: cell.fill = fill
                cell.border = _BRD
                cell.alignment = Alignment(vertical="center")

        def _total_row(ws, r, label, total):
            ws.merge_cells(f"A{r}:D{r}")
            c = ws[f"A{r}"]
            c.value = label
            c.font = Font(bold=True, size=10)
            c.fill = _LGRY
            c.border = _BRD
            c.alignment = Alignment(horizontal="right")
            cv = ws.cell(row=r, column=5, value=total)
            cv.font = Font(bold=True, size=10)
            cv.fill = _LGRY
            cv.border = _BRD
            cv.alignment = Alignment(horizontal="right")

        def build_excel(df_ext, df_inf, cheques, notas_cargo, notas_abono, ingresos,
                        saldo_ext, saldo_cont, nombre_pdf, matching_confiable=True):
            wb = Workbook()
            ws = wb.active; ws.title = "Conciliacion"
            for col, w in zip("ABCDE", [14, 40, 18, 18, 18]):
                ws.column_dimensions[col].width = w

            # ── Título ────────────────────────────────────────────────────────
            _merge_row(ws, 1, "CONCILIACIÓN BANCARIA",
                       PatternFill("solid", fgColor="1F3864"),
                       font_kw={"size":16,"bold":True,"color":"FFFFFF"},
                       height=32)
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

            _merge_row(ws, 2, f"{nombre_pdf}   |   {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                       PatternFill("solid", fgColor="F2F2F2"),
                       font_kw={"bold":False,"size":9,"color":"666666"})
            ws["A2"].alignment = Alignment(horizontal="center")

            # ── Saldos ────────────────────────────────────────────────────────
            r = 4
            diferencia = round((saldo_ext or 0) - (saldo_cont or 0), 2)
            _kv(ws, r,   "SALDO EXTRACTO BANCARIO",   _fmt_cop(saldo_ext),  _GRN); r += 1
            _kv(ws, r,   "SALDO CONTABILIDAD",        _fmt_cop(saldo_cont), _GRN); r += 1
            _kv(ws, r,   "DIFERENCIA A CONCILIAR",    _fmt_cop(diferencia),
                _GRN if diferencia == 0 else _YELL); r += 1

            if not matching_confiable:
                _merge_row(ws, r,
                    "⚠️  El extracto bancario agrupa pagos en lotes (dispersiones) y usa fechas valor distintas a "
                    "la contabilidad. El sistema no puede cruzar automáticamente las partidas individuales. "
                    "La diferencia a conciliar es la indicada arriba. Revisa las hojas 'Extracto' y 'Auxiliar' "
                    "para identificar el origen de la diferencia.",
                    PatternFill("solid", fgColor="FFF2CC"),
                    font_kw={"bold": False, "size": 9, "color": "7F4F00"})
                ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
                ws.row_dimensions[r].height = 52
                r += 2
            else:
                r += 1

            # ── Cheques pendientes de cobro ───────────────────────────────────
            _merge_row(ws, r, "CHEQUES PENDIENTES DE COBRO", _YELL); r += 1
            _hrow(ws, r, ["FECHA","BENEFICIARIO","No. COMPROBANTE","DOC","VALOR"]); r += 1
            if cheques.empty:
                _merge_row(ws, r, "—", PatternFill(), font_kw={"bold":False,"size":9,"color":"999999"}); r += 1
            else:
                for _, row in cheques.iterrows():
                    _drow(ws, r, [str(row.fecha), row.beneficiario, row.comprobante, "", _fmt_cop(row.monto)]); r += 1
            _total_row(ws, r, "TOTAL CHEQUES PENDIENTES POR COBRAR",
                       _fmt_cop(cheques.monto.sum()) if not cheques.empty else _fmt_cop(0)); r += 2

            # ── Disfones devueltos ────────────────────────────────────────────
            _merge_row(ws, r, "DISFONES DEVUELTOS", _YELL); r += 1
            _total_row(ws, r, "TOTAL DISFONES DEVUELTOS POR GIRAR", _fmt_cop(0)); r += 2

            # ── Notas no contabilizadas ───────────────────────────────────────
            _merge_row(ws, r, "NOTAS NO CONTABILIZADAS", _YELL); r += 1
            _hrow(ws, r, ["FECHA","DESCRIPCIÓN","CARGO","ABONO",""]); r += 1
            for _, row in notas_cargo.iterrows():
                _drow(ws, r, [str(row.fecha), row.descripcion, _fmt_cop(row.monto), "", ""]); r += 1
            for _, row in notas_abono.iterrows():
                _drow(ws, r, [str(row.fecha), row.descripcion, "", _fmt_cop(row.monto), ""]); r += 1
            if notas_cargo.empty and notas_abono.empty:
                _merge_row(ws, r, "—", PatternFill(), font_kw={"bold":False,"size":9,"color":"999999"}); r += 1
            total_notas = (notas_cargo.monto.sum() if not notas_cargo.empty else 0) + \
                          (notas_abono.monto.sum() if not notas_abono.empty else 0)
            _total_row(ws, r, "TOTAL NOTAS NO CONTABILIZADAS", _fmt_cop(total_notas)); r += 2

            # ── Ingresos pendientes por contabilizar ──────────────────────────
            _merge_row(ws, r, "INGRESOS PENDIENTES POR CONTABILIZAR", _YELL); r += 1
            _hrow(ws, r, ["FECHA","DESCRIPCIÓN","No. COMPROBANTE","","VALOR"]); r += 1
            if ingresos.empty:
                _merge_row(ws, r, "—", PatternFill(), font_kw={"bold":False,"size":9,"color":"999999"}); r += 1
            else:
                for _, row in ingresos.iterrows():
                    _drow(ws, r, [str(row.fecha), row.descripcion, row.comprobante, "", _fmt_cop(row.monto)]); r += 1
            _total_row(ws, r, "TOTAL INGRESOS PENDIENTES POR CONTABILIZAR",
                       _fmt_cop(ingresos.monto.sum()) if not ingresos.empty else _fmt_cop(0)); r += 2

            # ── Diferencia conciliada ─────────────────────────────────────────
            # Fórmula:
            # dif_conciliada = diferencia - cheques + ingresos + notas_cargo - notas_abono
            tc = cheques.monto.sum()  if not cheques.empty  else 0
            ti = ingresos.monto.sum() if not ingresos.empty else 0
            tnc = notas_cargo.monto.sum()  if not notas_cargo.empty  else 0
            tna = notas_abono.monto.sum()  if not notas_abono.empty  else 0
            dif_conciliada = round(diferencia - tc + ti + tnc - tna, 2)

            fill_conc = _GRN if abs(dif_conciliada) < 1 else _RED
            _kv(ws, r, "DIFERENCIA CONCILIADA", _fmt_cop(dif_conciliada), fill_conc)

            # ── Hoja 2: Extracto (detalle) ────────────────────────────────────
            ws2 = wb.create_sheet("Extracto")
            for col, w in zip("ABCDEF",[12,8,45,22,22,26]): ws2.column_dimensions[col].width = w
            _hrow(ws2, 1, ["Fecha","Día","Descripción","Débito","Crédito","Saldo"])
            for i,(_, row) in enumerate(df_ext.iterrows(), 2):
                fill = _GRN if i%2==0 else _GREY
                _drow(ws2, i, [str(row.fecha), row.dia, row.descripcion,
                    _fmt_cop(row.debito) if row.debito>0 else "",
                    _fmt_cop(row.credito) if row.credito>0 else "",
                    _fmt_cop(row.saldo)], fill)

            # ── Hoja 3: Auxiliar contable (detalle) ───────────────────────────
            ws3 = wb.create_sheet("Auxiliar")
            for col, w in zip("ABCDEF",[18,12,40,22,22,26]): ws3.column_dimensions[col].width = w
            _hrow(ws3, 1, ["Comprobante","Fecha","Descripción","Débito","Crédito","Saldo"])
            for i,(_, row) in enumerate(df_inf.iterrows(), 2):
                fill = _GRN if i%2==0 else _GREY
                _drow(ws3, i, [row.comprobante, str(row.fecha), row.descripcion,
                    _fmt_cop(row.debito) if row.debito>0 else "",
                    _fmt_cop(row.credito) if row.credito>0 else "",
                    _fmt_cop(row.saldo)], fill)


            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf, diferencia

        # ── Ejecutar ──────────────────────────────────────────────────────────
        meses_str = {"ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
                     "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12}
        # Intenta "01-MAR-2026", luego "MAR.2026" o "MAR 2026"
        m = re.search(r"(\d{1,2})[\.\-]([A-Za-z]{3,4})[\.\-](\d{4})", uploaded_ext.name)
        if m:
            ano = int(m.group(3))
            mes = meses_str.get(m.group(2).lower()[:3], 1)
        else:
            m2 = re.search(r"\b([A-Za-z]{3,4})[\.\-\s](\d{4})\b", uploaded_ext.name)
            if m2:
                ano = int(m2.group(2))
                mes = meses_str.get(m2.group(1).lower()[:3], date.today().month)
            else:
                ano, mes = date.today().year, date.today().month

        # Procesar extracto — pdfplumber lee texto directo, sin OCR
        df_ext, _images_ext = parse_extracto(ano, mes)
        _ext_saldos = df_ext["saldo"].dropna() if not df_ext.empty else pd.Series(dtype=float)
        saldo_ext = float(_ext_saldos.iloc[-1]) if not _ext_saldos.empty else None

        # Si pdfplumber no encontró saldo, intentar buscar en texto del PDF
        if saldo_ext is None and not df_ext.empty:
            try:
                with pdfplumber.open(pdf_ext_path) as pdf:
                    for page in reversed(pdf.pages):
                        txt = page.extract_text() or ""
                        for line in txt.split("\n"):
                            if re.search(r"saldo\s+(final|al\s+corte|bancario)", line, re.IGNORECASE):
                                nums = re.findall(r"[\d,]+\.\d{2}", line)
                                for n in nums:
                                    v = _to_float(n)
                                    if v and v > 0:
                                        saldo_ext = v; break
                        if saldo_ext: break
            except Exception:
                pass

        # Procesar auxiliar — pdfplumber lee texto directo, sin OCR
        df_inf, _images_aux = parse_auxiliar()
        _inf_saldos = df_inf["saldo"].dropna() if not df_inf.empty else pd.Series(dtype=float)
        saldo_cont = float(_inf_saldos.iloc[-1]) if not _inf_saldos.empty else None

        diferencia_saldos = round((saldo_ext or 0) - (saldo_cont or 0), 2)
        cheques, notas_cargo, notas_abono, ingresos, matching_ok = reconciliar(
            df_ext, df_inf, diferencia_saldos
        )

        excel_buf, diferencia = build_excel(
            df_ext, df_inf, cheques, notas_cargo, notas_abono, ingresos,
            saldo_ext, saldo_cont, uploaded_ext.name,
            matching_confiable=matching_ok,
        )

        os.unlink(pdf_ext_path)
        os.unlink(pdf_aux_path)

    # ── Resultado ─────────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("Saldo Extracto",      f"${saldo_ext:,.0f}"  if saldo_ext  else "—")
    col2.metric("Saldo Contabilidad",  f"${saldo_cont:,.0f}" if saldo_cont else "—")
    col3.metric("Diferencia",          f"${diferencia:,.0f}" if diferencia is not None else "—",
                delta_color="off" if diferencia==0 else "inverse")

    if diferencia == 0:
        st.success("✅ Conciliación completa — los saldos cuadran perfectamente.")
    elif not matching_ok:
        st.warning(
            f"⚠️ Diferencia de **${diferencia:,.0f}**. "
            "El extracto usa fechas valor y agrupa pagos en lotes (dispersiones), "
            "lo que impide el cruce automático. "
            "Las secciones del Excel están vacías — usa las hojas de detalle para identificar la partida."
        )
    else:
        st.error(f"⚠️ Diferencia de ${diferencia:,.0f} — revisa las partidas pendientes en el Excel.")

    nombre_salida = uploaded_ext.name.replace(".pdf", "_conciliacion.xlsx")
    st.download_button(
        label="⬇️ Descargar Excel de conciliación",
        data=excel_buf,
        file_name=nombre_salida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    with st.expander("Ver movimientos del extracto"):
        cols_ext = [c for c in ["fecha","dia","descripcion","debito","credito","saldo"] if c in df_ext.columns]
        if cols_ext:
            st.dataframe(df_ext[cols_ext], use_container_width=True)
        else:
            st.warning("pdfplumber no extrajo filas del extracto — ver texto de debug abajo.")
        debug = st.session_state.get("ext_debug", [])
        if debug:
            modo = debug[0][2] if debug else "?"
            st.caption(f"Modo: {'pdfplumber ✓' if modo=='pdf' else 'OCR'} — "
                       "Filas/página: " + " | ".join(f"P{p}:{n}" for p, n, _ in debug))
        sample = st.session_state.get("plumber_sample", [])
        if sample:
            st.caption("**Texto que leyó pdfplumber del extracto (pág 1) — cópiame esto:**")
            for ln in sample:
                st.code(ln)

    with st.expander("Ver asientos del contador"):
        cols_inf = [c for c in ["comprobante","fecha","descripcion","debito","credito","saldo"] if c in df_inf.columns]
        if cols_inf:
            st.dataframe(df_inf[cols_inf], use_container_width=True)
        else:
            st.warning("pdfplumber no extrajo filas del auxiliar.")
        aux_sample = st.session_state.get("aux_sample", [])
        if aux_sample:
            st.caption("**Texto que leyó pdfplumber del auxiliar (pág 1) — cópiame esto:**")
            for ln in aux_sample:
                st.code(ln)
