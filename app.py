"""
App Streamlit - Generador REP_2 (Costos y Gastos por Familia de Servicios)

Sube GRH_8.xlsx y GRH_11.xlsx (y opcionalmente la plantilla REP_2.xlsx con el
diccionario de atributos SISS) y la app arma la tabla REP_2 lista para
descargar, con validaciones de cuadratura incluidas.
"""

import io
import shutil
from collections import defaultdict

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Generador REP_2 - Suralis", layout="wide")

HEADERS = [
    "CÓDIGO EMPRESA", "PERÍODO INFORMACIÓN", "AÑO INFORMADO",
    "CÓDIGO SECTOR DECRETO TARIFARIO", "CÓDIGO RECURSO",
    "CÓDIGO FAMILIA SERVICIOS REGULADOS",
    "% NO ACTIVADO ASIGNADO FAMILIA SERVICIOS", "GASTO ANUAL",
    "% ACTIVADO ASIGNADO FAMILIA SERVICIOS", "MONTO ACTIVADO",
]


def read_rows(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    return list(ws.iter_rows(min_row=2, values_only=True))


def build_rep2(grh8_bytes, grh11_bytes):
    rows8 = read_rows(grh8_bytes)
    rows11 = read_rows(grh11_bytes)

    # Shares por (persona,cargo): TODOS los servicios (regulados y no regulados)
    shares = defaultdict(list)
    for r in rows11:
        cod_reg, cod_noreg, pct = r[7], r[8], r[9]
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        shares[(r[4], r[5])].append((cod_serv, pct))

    agg = defaultdict(lambda: [0.0, 0.0])
    for r in rows8:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r
        for cod_serv, pct in shares.get((persona, cargo), []):
            familia = cod_serv // 100
            k = (empresa, periodo, anio, sector, cod_recurso, familia)
            agg[k][0] += total_gasto * pct
            agg[k][1] += monto_act * pct

    tot_no_act_recurso = defaultdict(float)
    tot_act_recurso = defaultdict(float)
    for (empresa, periodo, anio, sector, cod_recurso, familia), (gasto, act) in agg.items():
        tot_no_act_recurso[cod_recurso] += gasto
        tot_act_recurso[cod_recurso] += act

    final_rows = []
    for (empresa, periodo, anio, sector, cod_recurso, familia), (gasto, act) in sorted(
            agg.items(), key=lambda x: (x[0][4], x[0][5])):
        if abs(gasto) < 1e-9 and abs(act) < 1e-9:
            continue
        pct_no_act = gasto / tot_no_act_recurso[cod_recurso] if tot_no_act_recurso[cod_recurso] else 0.0
        pct_act = act / tot_act_recurso[cod_recurso] if tot_act_recurso[cod_recurso] else 0.0
        final_rows.append([
            empresa, periodo, anio, sector, cod_recurso, familia,
            round(pct_no_act, 4) if gasto > 0 else 0.0,
            round(gasto, 2),
            round(pct_act, 4) if act > 0 else 0.0,
            round(act, 2),
        ])

    sum_gasto = sum(r[7] for r in final_rows)
    sum_grh8_total = sum(r[10] for r in rows8)

    sum_act = sum(r[9] for r in final_rows)
    sum_grh8_act = sum(r[8] for r in rows8)

    checks = {
        "sum_gasto_rep2": sum_gasto,
        "sum_gasto_grh8": sum_grh8_total,
        "diff_gasto": sum_gasto - sum_grh8_total,
        "sum_act_rep2": sum_act,
        "sum_act_grh8": sum_grh8_act,
        "diff_act": sum_act - sum_grh8_act,
    }
    return final_rows, checks


def build_excel(final_rows, template_bytes=None):
    if template_bytes is not None:
        tmp = io.BytesIO(template_bytes)
        wb = openpyxl.load_workbook(tmp)
        wb.active.title = "Diccionario_REP_2"
        ws = wb.create_sheet("REP_2", 0)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "REP_2"

    header_fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    data_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 30

    for row in final_rows:
        ws.append(row)

    for r in range(2, ws.max_row + 1):
        for c in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
            if c in (7, 9):
                cell.number_format = "0.00%"
            elif c in (8, 10):
                cell.number_format = '#,##0;(#,##0);"-"'
            else:
                cell.alignment = Alignment(horizontal="center")

    widths = [14, 16, 14, 20, 14, 22, 18, 18, 18, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------------- UI --------
st.title("Generador REP_2 · Costos y Gastos por Familia de Servicios")
st.caption(
    "Cruza GRH_8 (gasto por recurso) con GRH_11 (% de dedicación a servicios "
    "regulados) para construir la tabla REP_2 exigida por la SISS."
)

with st.sidebar:
    st.header("Archivos de entrada")
    f_grh8 = st.file_uploader("GRH_8.xlsx", type="xlsx")
    f_grh11 = st.file_uploader("GRH_11.xlsx", type="xlsx")
    f_template = st.file_uploader("REP_2.xlsx (plantilla/diccionario, opcional)", type="xlsx")
    run = st.button("Generar REP_2", type="primary", disabled=not (f_grh8 and f_grh11))

st.markdown(
    """
**Lógica aplicada**
1. GRH_8 entrega el gasto por (ID Persona, ID Cargo, Código Recurso).
2. GRH_11 entrega el % de dedicación de cada persona a **todos** sus servicios
   (regulados y no regulados) — cada fila trae un código en una de las dos
   columnas de servicio y su % asociado.
3. Ese % se aplica a los montos de GRH_8 (no activado y activado) de esa
   persona, ya que GRH_11 no abre por recurso.
4. Familia de Servicio = primeros 2 dígitos del código de servicio
   (regulado o no regulado, ej: 1101→11, 2201→22).
5. Se agrega por (Código Recurso, Familia) y se recalculan los % para que
   sumen 100% dentro de cada recurso, cubriendo el 100% del gasto de GRH_8.
    """
)

if run:
    try:
        final_rows, checks = build_rep2(f_grh8.getvalue(), f_grh11.getvalue())
    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
        st.stop()

    df = pd.DataFrame(final_rows, columns=HEADERS)

    st.success(f"REP_2 generado con {len(df)} filas.")

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Diferencia GASTO ANUAL vs GRH_8 (no activado)", f"{checks['diff_gasto']:,.2f}")
    with c2:
        st.metric("Diferencia MONTO ACTIVADO vs GRH_8", f"{checks['diff_act']:,.2f}")

    if abs(checks["diff_gasto"]) > 1 or abs(checks["diff_act"]) > 1:
        st.warning("Hay una diferencia de cuadratura mayor a $1. Revisa los archivos fuente.")
    else:
        st.info("Cuadratura OK: el 100% del gasto de GRH_8 (regulado + no regulado) quedó repartido en REP_2.")

    st.dataframe(
        df.style.format({
            "% NO ACTIVADO ASIGNADO FAMILIA SERVICIOS": "{:.2%}",
            "% ACTIVADO ASIGNADO FAMILIA SERVICIOS": "{:.2%}",
            "GASTO ANUAL": "{:,.0f}",
            "MONTO ACTIVADO": "{:,.0f}",
        }),
        use_container_width=True,
    )

    excel_bytes = build_excel(final_rows, f_template.getvalue() if f_template else None)
    st.download_button(
        "Descargar REP_2.xlsx",
        data=excel_bytes,
        file_name="REP_2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Sube GRH_8.xlsx y GRH_11.xlsx en el panel izquierdo y presiona **Generar REP_2**.")
