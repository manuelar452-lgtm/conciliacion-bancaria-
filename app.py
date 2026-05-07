from __future__ import annotations
import re
import io
import unicodedata
import numpy as np
import pandas as pd
from datetime import date, datetime, timedelta

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Conciliación Bancaria",
    page_icon="🏦",
    layout="centered",
)

col_a, col_b = st.columns(2)
uploaded_ext = col_a.file_uploader("📄 Extracto bancario (Excel)", type=["xlsx", "xls"])
uploaded_aux = col_b.file_uploader("📒 Auxiliar contabilidad (Excel)", type=["xlsx", "xls"])

if not uploaded_ext or not uploaded_aux:
    st.info("Sube los dos archivos Excel para continuar.")
    st.stop()

col_a.success(uploaded_ext.name)
col_b.success(uploaded_aux.name)

if st.button("Generar conciliación", type="primary"):

    with st.spinner("Procesando..."):

        # ── Utilidades ────────────────────────────────────────────────────────

        def _fmt_cop(v):
            if v is None or (isinstance(v, float) and np.isnan(v)): return ""
            try:    return f"$ {float(v):>20,.0f}"
            except: return str(v)

        def _to_num(v) -> float:
            if v is None: return 0.0
            try:
                if pd.isna(v): return 0.0
            except TypeError:
                pass
            try:    return float(v)
            except: return 0.0

        def _to_date(v):
            if v is None: return None
            try:
                if pd.isna(v): return None
            except TypeError:
                pass
            if isinstance(v, pd.Timestamp): return v.date()
            if isinstance(v, datetime):     return v.date()
            if isinstance(v, date):         return v
            try: return pd.to_datetime(str(v)).date()
            except: return None

        # ── Excel parsers ─────────────────────────────────────────────────────

        def _find_header_row(df_raw, keywords, max_rows=20):
            """Row index where the most keywords match (min 2 required)."""
            best_row, best_score = None, 0
            for r in range(min(max_rows, len(df_raw))):
                row_str = " ".join(str(v).upper() for v in df_raw.iloc[r] if pd.notna(v))
                score = sum(1 for kw in keywords if kw in row_str)
                if score > best_score:
                    best_score, best_row = score, r
            return best_row if best_score >= 2 else None

        def _norm(s):
            """Uppercase + strip accents for header comparison."""
            s = str(s).upper()
            return "".join(
                c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn"
            )

        def _col_idx(header_vals, keywords):
            """Column index: exact normalized match first, then partial match."""
            nkws = [_norm(kw) for kw in keywords]
            for nkw in nkws:            # exact pass
                for i, v in enumerate(header_vals):
                    if pd.notna(v) and _norm(v) == nkw:
                        return i
            for nkw in nkws:            # partial pass
                for i, v in enumerate(header_vals):
                    if pd.notna(v) and nkw in _norm(v):
                        return i
            return None

        def _col_idx_last(header_vals, keywords):
            """Last column index whose normalized header contains any keyword."""
            result = None
            for kw in keywords:
                nkw = _norm(kw)
                for i, v in enumerate(header_vals):
                    if pd.notna(v) and nkw in _norm(v):
                        result = i
            return result

        def parse_extracto_excel(file_bytes, filename):
            engine = "xlrd" if filename.lower().endswith(".xls") else "openpyxl"
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None,
                                   engine=engine, dtype=object)

            # ── Saldo final ────────────────────────────────────────────────
            saldo_ext = None
            for r in range(min(10, len(df_raw))):
                row_vals = df_raw.iloc[r].tolist()
                row_str  = " ".join(str(v).upper() for v in row_vals if pd.notna(v))
                if "SALDO FINAL" in row_str or "SALDO AL CORTE" in row_str:
                    col_sf = next((c for c, v in enumerate(row_vals)
                                   if pd.notna(v) and "SALDO FINAL" in str(v).upper()), None)
                    if col_sf is not None and r + 1 < len(df_raw):
                        try:
                            fv = float(df_raw.iloc[r + 1, col_sf])
                            if fv > 100_000:
                                saldo_ext = fv
                        except Exception:
                            pass
                    break

            # ── Header row ──────────────────────────────────────────────────
            hdr_row = _find_header_row(
                df_raw,
                ["FECHA", "DESCRIPCI", "DEP", "RETIRO"],
                max_rows=15,
            )
            if hdr_row is None:
                st.warning("Extracto: no se encontró fila de encabezado.")
                return (pd.DataFrame(columns=["fecha","dia","descripcion",
                                               "debito","credito","saldo"]),
                        saldo_ext)

            hdr     = df_raw.iloc[hdr_row].tolist()
            col_f   = _col_idx(hdr, ["FECHA"])
            col_d   = _col_idx(hdr, ["DESCRIPCION"])
            col_cre = _col_idx(hdr, ["DEP", "ABONO", "CRÉDITO", "CREDITO"])
            col_deb = _col_idx(hdr, ["RETIRO", "CARGO", "DÉBITO", "DEBITO", "PAGO"])
            col_sal = _col_idx(hdr, ["SALDO"])

            st.session_state["ext_cols"] = (
                f"header_row={hdr_row} | fecha={col_f} | desc={col_d} | "
                f"dep(cre)={col_cre} | retiro(deb)={col_deb} | saldo={col_sal}"
            )

            rows = []
            for r in range(hdr_row + 1, len(df_raw)):
                row   = df_raw.iloc[r]
                fecha = _to_date(row.iloc[col_f] if col_f is not None else None)
                if fecha is None:
                    continue

                debito  = abs(_to_num(row.iloc[col_deb] if col_deb is not None else None))
                credito = abs(_to_num(row.iloc[col_cre] if col_cre is not None else None))
                if debito == 0 and credito == 0:
                    continue

                saldo = abs(_to_num(row.iloc[col_sal] if col_sal is not None else None))
                desc  = (str(row.iloc[col_d]).strip()
                         if col_d is not None and pd.notna(row.iloc[col_d]) else "")

                rows.append({
                    "fecha": fecha, "dia": fecha.day,
                    "descripcion": desc,
                    "debito": debito, "credito": credito, "saldo": saldo,
                })

            if not rows:
                return (pd.DataFrame(columns=["fecha","dia","descripcion",
                                               "debito","credito","saldo"]),
                        saldo_ext)

            df = pd.DataFrame(rows).sort_values("fecha").reset_index(drop=True)
            return df, saldo_ext

        def parse_auxiliar_excel(file_bytes, filename):
            engine = "xlrd" if filename.lower().endswith(".xls") else "openpyxl"
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None,
                                   engine=engine, dtype=object)

            # ── Header row ──────────────────────────────────────────────────
            hdr_row = _find_header_row(
                df_raw,
                ["COMPROBANTE", "FECHA", "DEBITO", "CREDITO"],
                max_rows=20,
            )
            if hdr_row is None:
                st.warning("Auxiliar: no se encontró fila de encabezado.")
                return (pd.DataFrame(columns=["comprobante","fecha","descripcion",
                                               "debito","credito","saldo"]),
                        None)

            hdr      = df_raw.iloc[hdr_row].tolist()
            col_comp = _col_idx(hdr, ["COMPROBANTE"])
            col_f    = _col_idx(hdr, ["FECHA"])
            col_nom  = _col_idx(hdr, ["NOMBRE"])
            col_desc = _col_idx(hdr, ["DESCRIPCION"])
            col_deb  = _col_idx(hdr, ["DEBITO", "DÉBITO"])
            col_cre  = _col_idx(hdr, ["CREDITO", "CRÉDITO"])
            # SALDO MOV. is the rightmost saldo; "SALDO CUENTA" is leftmost — take last match
            col_sal  = _col_idx(hdr, ["SALDO MOV"])
            if col_sal is None:
                col_sal = _col_idx_last(hdr, ["SALDO"])

            st.session_state["aux_cols"] = (
                f"header_row={hdr_row} | comp={col_comp} | fecha={col_f} | "
                f"nombre={col_nom} | desc={col_desc} | "
                f"deb={col_deb} | cre={col_cre} | saldo={col_sal}"
            )

            _SKIP = {"TOTAL", "SUBTOTAL", "SALDO", "CUENTA", "NAN", "NONE"}

            rows = []
            for r in range(hdr_row + 1, len(df_raw)):
                row  = df_raw.iloc[r]
                comp = (str(row.iloc[col_comp]).strip()
                        if col_comp is not None and pd.notna(row.iloc[col_comp])
                        else "")
                if not comp or any(kw in comp.upper() for kw in _SKIP):
                    continue

                fecha = _to_date(row.iloc[col_f] if col_f is not None else None)
                if fecha is None:
                    continue

                debito  = abs(_to_num(row.iloc[col_deb] if col_deb is not None else None))
                credito = abs(_to_num(row.iloc[col_cre] if col_cre is not None else None))
                if debito == 0 and credito == 0:
                    continue

                saldo = _to_num(row.iloc[col_sal] if col_sal is not None else None)

                parts = []
                if col_nom is not None and pd.notna(row.iloc[col_nom]) and str(row.iloc[col_nom]).strip():
                    parts.append(str(row.iloc[col_nom]).strip())
                if col_desc is not None and pd.notna(row.iloc[col_desc]) and str(row.iloc[col_desc]).strip():
                    parts.append(str(row.iloc[col_desc]).strip())
                desc = " - ".join(parts) if parts else comp

                rows.append({
                    "comprobante": comp, "fecha": fecha,
                    "descripcion": desc,
                    "debito": debito, "credito": credito, "saldo": saldo,
                })

            if not rows:
                return (pd.DataFrame(columns=["comprobante","fecha","descripcion",
                                               "debito","credito","saldo"]),
                        None)

            df = pd.DataFrame(rows).sort_values("fecha").reset_index(drop=True)
            return df, None

        # ── Matching ──────────────────────────────────────────────────────────
        def reconciliar(df_ext, df_inf, diferencia_saldos):
            ext_idx = {}
            for i, row in df_ext.iterrows():
                for monto, tipo in [(row.debito, "D"), (row.credito, "C")]:
                    if monto > 0:
                        key = (row.fecha, round(monto / 100) * 100, tipo)
                        ext_idx.setdefault(key, []).append(i)

            matched_ext, matched_inf = set(), set()
            for i, row in df_inf.iterrows():
                for monto, tipo_ext in [(row.debito, "C"), (row.credito, "D")]:
                    if monto <= 0: continue
                    found    = False
                    tol_amt  = max(round(monto * 0.01 / 100) * 100, 1000)
                    for day_d in [0,-1,-2,-3,-4,-5,-6,-7,-8,-9,-10,1,2,3,4,5]:
                        fecha_try = row.fecha + timedelta(days=day_d)
                        for amt_step in range(0, int(tol_amt) + 100, 100):
                            for sign in ([0] if amt_step == 0 else [amt_step, -amt_step]):
                                key = (fecha_try, round((monto + sign) / 100) * 100, tipo_ext)
                                if key in ext_idx and ext_idx[key]:
                                    matched_ext.add(ext_idx[key].pop(0))
                                    matched_inf.add(i)
                                    found = True; break
                            if found: break
                        if found: break

            _e   = lambda lst, cols: pd.DataFrame(lst) if lst else pd.DataFrame(columns=cols)
            emp  = lambda cols: pd.DataFrame(columns=cols)

            if round(abs(diferencia_saldos)) == 0:
                return (emp(["fecha","beneficiario","comprobante","monto"]),
                        emp(["fecha","descripcion","monto"]),
                        emp(["fecha","descripcion","monto"]),
                        emp(["fecha","descripcion","comprobante","monto"]),
                        True)

            notas_cargo, notas_abono = [], []
            for i, row in df_ext.iterrows():
                if i in matched_ext: continue
                if row.debito > 0:
                    notas_cargo.append({"fecha": row.fecha, "descripcion": row.descripcion,
                                        "monto": row.debito})
                if row.credito > 0:
                    notas_abono.append({"fecha": row.fecha, "descripcion": row.descripcion,
                                        "monto": row.credito})

            cheques_pend, ingresos_pend = [], []
            for i, row in df_inf.iterrows():
                if i in matched_inf: continue
                if row.credito > 0:
                    cheques_pend.append({"fecha": row.fecha, "beneficiario": row.descripcion,
                                         "comprobante": row.comprobante, "monto": row.credito})
                if row.debito > 0:
                    ingresos_pend.append({"fecha": row.fecha, "descripcion": row.descripcion,
                                          "comprobante": row.comprobante, "monto": row.debito})

            notas_cargo.sort(  key=lambda x: x["monto"], reverse=True)
            notas_abono.sort(  key=lambda x: x["monto"], reverse=True)
            cheques_pend.sort( key=lambda x: x["monto"], reverse=True)
            ingresos_pend.sort(key=lambda x: x["monto"], reverse=True)

            return (
                _e(cheques_pend,  ["fecha","beneficiario","comprobante","monto"]),
                _e(notas_cargo,   ["fecha","descripcion","monto"]),
                _e(notas_abono,   ["fecha","descripcion","monto"]),
                _e(ingresos_pend, ["fecha","descripcion","comprobante","monto"]),
                True,
            )

        def _buscar_partida(df_ext, df_inf, diferencia_saldos):
            """
            Devuelve lista de partidas (greedy) que juntas explican la diferencia.
            Busca en AMBAS direcciones:
              - diferencia < 0 (aux > banco): créditos en aux sin retiro en banco
                                             + débitos en banco sin registro en aux
              - diferencia > 0 (banco > aux): depósitos en banco sin registro en aux
                                             + débitos en aux sin depósito en banco
            """
            from collections import defaultdict

            if df_ext.empty or df_inf.empty:
                return []

            target = abs(diferencia_saldos)
            if target < 1_000:
                return []

            def _key(v):
                return round(v / 1_000) * 1_000

            def _pool_from(series):
                p = defaultdict(int)
                for v in series:
                    if v > 0: p[_key(v)] += 1
                return p

            def _consume(pool, monto, tol=5_000):
                mk = _key(monto)
                best_k, best_d = None, float("inf")
                for k, cnt in pool.items():
                    if cnt > 0 and abs(k - mk) <= tol and abs(k - mk) < best_d:
                        best_d, best_k = abs(k - mk), k
                if best_k is not None:
                    pool[best_k] -= 1
                    return True
                return False

            todos = []

            if diferencia_saldos < 0:
                # aux > banco → buscar partidas que INFLAN saldo_cont o DEFLATAN saldo_ext
                # A) créditos en aux (pagos asentados) sin retiro correspondiente en banco
                ext_deb_pool = _pool_from(df_ext["debito"])
                for _, row in df_inf.iterrows():
                    monto = row.credito
                    if monto < 1_000: continue
                    if not _consume(ext_deb_pool, monto):
                        todos.append({
                            "origen": "Pago en auxiliar sin retiro en banco",
                            "fecha": row.fecha,
                            "descripcion": f"{row.comprobante} {row.descripcion}".strip(),
                            "monto": monto, "tipo": "CRÉDITO aux",
                        })
                # B) depósitos en banco sin débito en aux (banco recibió algo no asentado)
                aux_deb_pool = _pool_from(df_inf["debito"])
                for _, row in df_ext.iterrows():
                    monto = row.credito
                    if monto < 1_000: continue
                    if not _consume(aux_deb_pool, monto):
                        todos.append({
                            "origen": "Depósito en banco sin registro en auxiliar",
                            "fecha": row.fecha,
                            "descripcion": row.descripcion,
                            "monto": monto, "tipo": "CRÉDITO banco",
                        })
            else:
                # banco > aux → buscar partidas que INFLAN saldo_ext o DEFLATAN saldo_cont
                # A) depósitos en banco sin débito en aux
                aux_deb_pool = _pool_from(df_inf["debito"])
                for _, row in df_ext.iterrows():
                    monto = row.credito
                    if monto < 1_000: continue
                    if not _consume(aux_deb_pool, monto):
                        todos.append({
                            "origen": "Depósito en banco sin registro en auxiliar",
                            "fecha": row.fecha,
                            "descripcion": row.descripcion,
                            "monto": monto, "tipo": "CRÉDITO banco",
                        })
                # B) créditos en aux sin retiro en banco
                ext_deb_pool = _pool_from(df_ext["debito"])
                for _, row in df_inf.iterrows():
                    monto = row.credito
                    if monto < 1_000: continue
                    if not _consume(ext_deb_pool, monto):
                        todos.append({
                            "origen": "Pago en auxiliar sin retiro en banco",
                            "fecha": row.fecha,
                            "descripcion": f"{row.comprobante} {row.descripcion}".strip(),
                            "monto": monto, "tipo": "CRÉDITO aux",
                        })

            if not todos:
                return []

            # ── Greedy: seleccionar las que juntas cubren la diferencia ──────
            todos.sort(key=lambda x: x["monto"], reverse=True)
            seleccionadas = []
            restante = target
            pool_greedy = list(todos)

            while restante > target * 0.005 and pool_greedy:
                # Elegir la entrada cuyo monto esté más cerca del restante
                candidatas = [c for c in pool_greedy if c["monto"] <= restante * 1.5]
                if not candidatas:
                    break
                mejor = min(candidatas, key=lambda x: abs(x["monto"] - restante))
                seleccionadas.append(mejor)
                restante -= mejor["monto"]
                pool_greedy.remove(mejor)
                if len(seleccionadas) >= 15:
                    break

            if not seleccionadas:
                return []

            acum = 0
            for c in seleccionadas:
                acum += c["monto"]
                c["acumulado"] = acum
                c["restante"]  = max(0.0, target - acum)
                c["cobertura"] = round(acum / target * 100, 1) if target else 0

            return seleccionadas

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
            c.fill  = fill
            c.alignment = Alignment(horizontal="left", vertical="center")
            kw = {"bold": True, "size": 11, "color": "1F3864"}
            if font_kw: kw.update(font_kw)
            c.font = Font(**kw)
            if height: ws.row_dimensions[r].height = height

        def _kv(ws, r, label, valor, fill_v=None):
            ca = ws.cell(row=r, column=1, value=label)
            ca.font   = Font(bold=True, size=10)
            ca.border = _BRD
            cv = ws.cell(row=r, column=2, value=valor)
            cv.border    = _BRD
            cv.alignment = Alignment(horizontal="right", vertical="center")
            cv.font      = Font(size=10)
            if fill_v: cv.fill = fill_v

        def _hrow(ws, r, cols):
            for c, v in enumerate(cols, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font      = Font(bold=True, color="FFFFFF", size=9)
                cell.fill      = _HDR
                cell.alignment = Alignment(horizontal="center", vertical="center",
                                           wrap_text=True)
                cell.border = _BRD

        def _drow(ws, r, vals, fill=None):
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                if fill: cell.fill = fill
                cell.border    = _BRD
                cell.alignment = Alignment(vertical="center")

        def build_excel(df_ext, df_inf, saldo_ext, saldo_cont, nombre_arch, partidas):
            wb = Workbook()
            ws = wb.active; ws.title = "Conciliacion"
            for col, cw in zip("ABCDE", [14, 40, 22, 22, 22]):
                ws.column_dimensions[col].width = cw

            diferencia = round((saldo_ext or 0) - (saldo_cont or 0), 2)

            ws.merge_cells("A1:E1")
            c = ws["A1"]
            c.value     = "CONCILIACIÓN BANCARIA"
            c.font      = Font(bold=True, size=14, color="FFFFFF")
            c.fill      = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 30

            ws.merge_cells("A2:E2")
            c2       = ws["A2"]
            c2.value = f"{nombre_arch}   |   {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            c2.font  = Font(size=9, color="666666")
            c2.fill  = PatternFill("solid", fgColor="F2F2F2")
            c2.alignment = Alignment(horizontal="center")

            _kv(ws, 4, "SALDO EXTRACTO BANCARIO", _fmt_cop(saldo_ext),  _GRN)
            _kv(ws, 5, "SALDO CONTABILIDAD",       _fmt_cop(saldo_cont), _GRN)
            _kv(ws, 6, "DIFERENCIA",               _fmt_cop(diferencia),
                _GRN if diferencia == 0 else _RED)

            next_row = 8
            if partidas:
                total_cubierto = sum(p["monto"] for p in partidas)
                pct = round(total_cubierto / abs(diferencia) * 100, 1) if diferencia else 0

                ws.merge_cells(f"A{next_row}:E{next_row}")
                lbl       = ws[f"A{next_row}"]
                lbl.value = f"PARTIDAS EN CONCILIACIÓN  —  cubre {pct}% de la diferencia"
                lbl.font  = Font(bold=True, size=11, color="7F4F00")
                lbl.fill  = _YELL
                lbl.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[next_row].height = 22
                next_row += 1

                _hrow(ws, next_row,
                      ["Fecha", "Descripción / Origen", "Monto", "Acumulado", "Restante"])
                next_row += 1

                for p in partidas:
                    _drow(ws, next_row,
                          [str(p["fecha"]),
                           f"{p['descripcion']}  [{p['origen']}]",
                           _fmt_cop(p["monto"]),
                           _fmt_cop(p["acumulado"]),
                           _fmt_cop(p["restante"])],
                          _YELL)
                    next_row += 1

            ws2 = wb.create_sheet("Extracto")
            for col, cw in zip("ABCDEF", [12, 8, 45, 22, 22, 26]):
                ws2.column_dimensions[col].width = cw
            _hrow(ws2, 1, ["Fecha","Día","Descripción","Débito","Crédito","Saldo"])
            ext_rows = list(df_ext.iterrows())
            for i, (_, row) in enumerate(ext_rows, 2):
                fill     = _GRN if i % 2 == 0 else _GREY
                es_ult   = (i == len(ext_rows) + 1)
                saldo_f  = saldo_ext if (es_ult and saldo_ext) else row.saldo
                _drow(ws2, i,
                      [str(row.fecha), row.dia, row.descripcion,
                       _fmt_cop(row.debito)  if row.debito  > 0 else "",
                       _fmt_cop(row.credito) if row.credito > 0 else "",
                       _fmt_cop(saldo_f)],
                      fill)

            ws3 = wb.create_sheet("Auxiliar")
            for col, cw in zip("ABCDEF", [18, 12, 40, 22, 22, 26]):
                ws3.column_dimensions[col].width = cw
            _hrow(ws3, 1, ["Comprobante","Fecha","Descripción","Débito","Crédito","Saldo"])
            for i, (_, row) in enumerate(df_inf.iterrows(), 2):
                fill = _GRN if i % 2 == 0 else _GREY
                _drow(ws3, i,
                      [row.comprobante, str(row.fecha), row.descripcion,
                       _fmt_cop(row.debito)  if row.debito  > 0 else "",
                       _fmt_cop(row.credito) if row.credito > 0 else "",
                       _fmt_cop(row.saldo)],
                      fill)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf, diferencia

        # ── Ejecutar ──────────────────────────────────────────────────────────
        ext_bytes = uploaded_ext.read()
        aux_bytes = uploaded_aux.read()

        df_ext, saldo_ext = parse_extracto_excel(ext_bytes, uploaded_ext.name)

        df_inf, _ = parse_auxiliar_excel(aux_bytes, uploaded_aux.name)
        _inf_saldos = df_inf["saldo"].dropna() if not df_inf.empty else pd.Series(dtype=float)
        saldo_cont  = float(_inf_saldos.iloc[-1]) if not _inf_saldos.empty else None

        diferencia_saldos = round((saldo_ext or 0) - (saldo_cont or 0), 2)
        partidas          = _buscar_partida(df_ext, df_inf, diferencia_saldos)

        excel_buf, diferencia = build_excel(
            df_ext, df_inf, saldo_ext, saldo_cont,
            uploaded_ext.name, partidas,
        )

    # ── Resultado ─────────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("Saldo Extracto",     f"${saldo_ext:,.0f}"  if saldo_ext  else "—")
    col2.metric("Saldo Contabilidad", f"${saldo_cont:,.0f}" if saldo_cont else "—")
    col3.metric("Diferencia",         f"${diferencia:,.0f}" if diferencia is not None else "—",
                delta_color="off" if diferencia == 0 else "inverse")

    n_ext = len(df_ext) if not df_ext.empty else 0
    n_aux = len(df_inf) if not df_inf.empty else 0
    st.caption(f"Extracto: **{n_ext} movimientos**   |   Auxiliar: **{n_aux} movimientos**")

    if diferencia == 0:
        st.success("✅ Conciliación completa — los saldos cuadran perfectamente.")
    else:
        st.error(f"⚠️ Diferencia de **${diferencia:,.0f}**")

        if partidas:
            total_cubierto = sum(p["monto"] for p in partidas)
            pct = round(total_cubierto / abs(diferencia) * 100, 1) if diferencia else 0
            restante_total = abs(diferencia) - total_cubierto

            st.warning(
                f"**Partidas en conciliación encontradas: {len(partidas)}**  —  "
                f"cubren **{pct}%** de la diferencia  "
                f"(acumulado **${total_cubierto:,.0f}** | sin explicar: **${restante_total:,.0f}**)"
            )

            df_partidas = pd.DataFrame([{
                "Fecha":        p["fecha"],
                "Descripción":  p["descripcion"],
                "Origen":       p["origen"],
                "Monto":        p["monto"],
                "Acumulado":    p["acumulado"],
                "Restante":     p["restante"],
                "Cobertura %":  p["cobertura"],
            } for p in partidas])

            st.dataframe(
                df_partidas.style.format({
                    "Monto":       "${:,.0f}",
                    "Acumulado":   "${:,.0f}",
                    "Restante":    "${:,.0f}",
                    "Cobertura %": "{:.1f}%",
                }),
                use_container_width=True,
                hide_index=True,
            )

    nombre_salida = re.sub(r"\.(xlsx|xls)$", "", uploaded_ext.name,
                           flags=re.IGNORECASE) + "_conciliacion.xlsx"
    st.download_button(
        label="⬇️ Descargar Excel de conciliación",
        data=excel_buf,
        file_name=nombre_salida,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    with st.expander("Ver movimientos del extracto"):
        cols_ext = [c for c in ["fecha","dia","descripcion","debito","credito","saldo"]
                    if c in df_ext.columns]
        if cols_ext:
            st.dataframe(df_ext[cols_ext], use_container_width=True)
        else:
            st.warning("No se extrajeron filas del extracto.")
        dbg = st.session_state.get("ext_cols", "")
        if dbg:
            st.caption(f"Columnas detectadas: `{dbg}`")

    with st.expander("Ver asientos del auxiliar"):
        cols_inf = [c for c in ["comprobante","fecha","descripcion","debito","credito","saldo"]
                    if c in df_inf.columns]
        if cols_inf:
            st.dataframe(df_inf[cols_inf], use_container_width=True)
        else:
            st.warning("No se extrajeron filas del auxiliar.")
        dbg2 = st.session_state.get("aux_cols", "")
        if dbg2:
            st.caption(f"Columnas detectadas: `{dbg2}`")
