"""
App Streamlit - Generador REP_2 (Costos y Gastos por Familia de Servicios)

Consolida 3 familias de gasto:
  - GRH: Gastos de Recursos Humanos           (GRH_8 + GRH_11)
  - GCP: Gastos Generales de Personal         (GCP_4 + GCP_5, ya abierta por recurso)
  - GGV: Gastos Generales Vehículos y Equipos (GGV_4 + GGV_5)

Sube los archivos correspondientes y la app arma la tabla REP_2 consolidada,
lista para descargar, con validaciones de cuadratura por cada familia.
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


def familia(cod_servicio):
    return cod_servicio // 100


def build_rep2(grh8_bytes, grh11_bytes, gcp4_bytes, gcp5_bytes, ggv4_bytes, ggv5_bytes):
    grh8 = read_rows(grh8_bytes)
    grh11 = read_rows(grh11_bytes)
    gcp4 = read_rows(gcp4_bytes)
    gcp5 = read_rows(gcp5_bytes)
    ggv4 = read_rows(ggv4_bytes)
    ggv5 = read_rows(ggv5_bytes)

    agg = defaultdict(lambda: [0.0, 0.0])

    # --- GRH: % de dedicación por (persona,cargo) desde GRH_11 -------------
    shares_persona = defaultdict(list)
    for r in grh11:
        cod_reg, cod_noreg, pct = r[7], r[8], r[9]
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        shares_persona[(r[4], r[5])].append((cod_serv, pct))

    for r in grh8:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r
        for cod_serv, pct in shares_persona.get((persona, cargo), []):
            fam = familia(cod_serv)
            k = (empresa, periodo, anio, sector, cod_recurso, fam)
            agg[k][0] += total_gasto * pct
            agg[k][1] += monto_act * pct

    # --- GCP: GCP_5 ya viene abierta por recurso y servicio -----------------
    # Gasto NO activado: directo desde GCP_5 (ya es el monto asignado)
    for r in gcp5:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        fam = familia(cod_serv)
        k = (empresa, periodo, anio, sector, cod_recurso, fam)
        agg[k][0] += gasto_no_act

    # Monto activado: usa el % de GCP_5 (persona,cargo,recurso) aplicado a GCP_4
    shares_gcp_recurso = defaultdict(list)
    for r in gcp5:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        shares_gcp_recurso[(persona, cargo, cod_recurso)].append((cod_serv, pct))

    for r in gcp4:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r
        if monto_act == 0:
            continue
        for cod_serv, pct in shares_gcp_recurso.get((persona, cargo, cod_recurso), []):
            fam = familia(cod_serv)
            k = (empresa, periodo, anio, sector, cod_recurso, fam)
            agg[k][1] += monto_act * pct

    # --- GGV: % de dedicación por ID Activo desde GGV_5 ---------------------
    shares_activo = defaultdict(list)
    for r in ggv5:
        empresa, periodo, anio, sector, id_activo, total_no_act, cod_reg, cod_noreg, pct = r[:9]
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        shares_activo[id_activo].append((cod_serv, pct))

    for r in ggv4:
        empresa, periodo, anio, sector, id_activo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r[:10]
        for cod_serv, pct in shares_activo.get(id_activo, []):
            fam = familia(cod_serv)
            k = (empresa, periodo, anio, sector, cod_recurso, fam)
            agg[k][0] += total_gasto * pct
            agg[k][1] += monto_act * pct

    # --- Normalizar % por recurso -------------------------------------------
    tot_no_act = defaultdict(float)
    tot_act = defaultdict(float)
    for (empresa, periodo, anio, sector, cod_recurso, fam), (g, a) in agg.items():
        tot_no_act[cod_recurso] += g
        tot_act[cod_recurso] += a

    final_rows = []
    for (empresa, periodo, anio, sector, cod_recurso, fam), (g, a) in sorted(
            agg.items(), key=lambda x: (x[0][4], x[0][5])):
        if abs(g) < 1e-9 and abs(a) < 1e-9:
            continue
        pct_no_act = g / tot_no_act[cod_recurso] if tot_no_act[cod_recurso] else 0.0
        pct_act = a / tot_act[cod_recurso] if tot_act[cod_recurso] else 0.0
        final_rows.append([
            empresa, periodo, anio, sector, cod_recurso, fam,
            round(pct_no_act, 4) if g > 0 else 0.0, round(g, 2),
            round(pct_act, 4) if a > 0 else 0.0, round(a, 2),
        ])

    # --- Validaciones por familia --------------------------------------------
    recursos_grh = set(r[6] for r in grh8)
    recursos_gcp = set(r[6] for r in gcp4)
    recursos_ggv = set(r[5] for r in ggv4)

    def validar(recursos, sum_gasto_fuente, sum_act_fuente):
        sum_gasto_rep2 = sum(r[7] for r in final_rows if r[4] in recursos)
        sum_act_rep2 = sum(r[9] for r in final_rows if r[4] in recursos)
        return {
            "gasto_rep2": sum_gasto_rep2, "gasto_fuente": sum_gasto_fuente,
            "diff_gasto": sum_gasto_rep2 - sum_gasto_fuente,
            "act_rep2": sum_act_rep2, "act_fuente": sum_act_fuente,
            "diff_act": sum_act_rep2 - sum_act_fuente,
        }

    checks = {
        "GRH": validar(recursos_grh, sum(r[10] for r in grh8), sum(r[8] for r in grh8)),
        "GCP": validar(recursos_gcp, sum(r[10] for r in gcp4), sum(r[8] for r in gcp4)),
        "GGV": validar(recursos_ggv, sum(r[9] for r in ggv4), sum(r[7] for r in ggv4)),
    }
    familia_map = {}
    for c in recursos_grh:
        familia_map[c] = "GRH - Gastos Recursos Humanos"
    for c in recursos_gcp:
        familia_map[c] = "GCP - Gastos Generales de Personal"
    for c in recursos_ggv:
        familia_map[c] = "GGV - Gastos Generales Vehículos y Equipos"

    return final_rows, checks, familia_map


def build_excel(final_rows, familia_map, template_bytes=None):
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

    # Hoja resumen por familia de gasto
    ws2 = wb.create_sheet("Resumen_por_familia_gasto")
    ws2.append(["Familia de Gasto", "GASTO ANUAL (no activado)", "MONTO ACTIVADO", "TOTAL"])
    for c in range(1, 5):
        cell = ws2.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    resumen = defaultdict(lambda: [0.0, 0.0])
    for row in final_rows:
        fg = familia_map.get(row[4], "Otro")
        resumen[fg][0] += row[7]
        resumen[fg][1] += row[9]

    for fg, (g, a) in resumen.items():
        ws2.append([fg, round(g, 2), round(a, 2), round(g + a, 2)])

    for r in range(2, ws2.max_row + 1):
        for c in range(1, 5):
            cell = ws2.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
            if c in (2, 3, 4):
                cell.number_format = '#,##0;(#,##0);"-"'
    ws2.column_dimensions["A"].width = 42
    for col in ["B", "C", "D"]:
        ws2.column_dimensions[col].width = 20

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------------- UI --------
st.title("Generador REP_2 · Costos y Gastos por Familia de Servicios")
st.caption(
    "Consolida las familias de gasto GRH (Recursos Humanos), GCP (Gastos "
    "Generales de Personal) y GGV (Vehículos y Equipos) en la tabla REP_2 "
    "exigida por la SISS."
)

with st.sidebar:
    st.header("Archivos de entrada")
    st.markdown("**GRH — Recursos Humanos**")
    f_grh8 = st.file_uploader("GRH_8.xlsx", type="xlsx", key="grh8")
    f_grh11 = st.file_uploader("GRH_11.xlsx", type="xlsx", key="grh11")
    st.markdown("**GCP — Gastos Generales de Personal**")
    f_gcp4 = st.file_uploader("GCP_4.xlsx", type="xlsx", key="gcp4")
    f_gcp5 = st.file_uploader("GCP_5.xlsx", type="xlsx", key="gcp5")
    st.markdown("**GGV — Vehículos y Equipos**")
    f_ggv4 = st.file_uploader("GGV_4.xlsx", type="xlsx", key="ggv4")
    f_ggv5 = st.file_uploader("GGV_5.xlsx", type="xlsx", key="ggv5")
    st.markdown("**Plantilla (opcional)**")
    f_template = st.file_uploader("REP_2.xlsx (diccionario)", type="xlsx", key="template")

    all_ready = all([f_grh8, f_grh11, f_gcp4, f_gcp5, f_ggv4, f_ggv5])
    run = st.button("Generar REP_2", type="primary", disabled=not all_ready)

st.markdown(
    """
**Lógica aplicada**
- **GRH**: el % de dedicación de cada persona a cada servicio (GRH_11, a
  nivel persona, filas regulado + no regulado) se aplica a los montos de
  GRH_8.
- **GCP**: GCP_5 ya viene abierta por (persona, cargo, **recurso**) y
  servicio — el gasto no activado se toma directo de ahí. El monto activado
  (de GCP_4) se reparte usando el mismo % de GCP_5, ahora a nivel de recurso.
- **GGV**: el % de dedicación de cada activo (GGV_5) se aplica a los montos
  de GGV_4, misma lógica que GRH pero a nivel de ID Activo.
- Familia de Servicio = primeros 2 dígitos del código de servicio.
- Se agrega todo por (Código Recurso, Familia) y se recalculan los % para
  que sumen 100% dentro de cada recurso.
    """
)

if run:
    try:
        final_rows, checks, familia_map = build_rep2(
            f_grh8.getvalue(), f_grh11.getvalue(),
            f_gcp4.getvalue(), f_gcp5.getvalue(),
            f_ggv4.getvalue(), f_ggv5.getvalue()
        )
    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
        st.stop()

    df = pd.DataFrame(final_rows, columns=HEADERS)
    st.success(f"REP_2 generado con {len(df)} filas (3 familias de gasto consolidadas).")

    st.subheader("Validación de cuadratura por familia de gasto")
    cols = st.columns(3)
    all_ok = True
    for col, (fam, chk) in zip(cols, checks.items()):
        with col:
            st.markdown(f"**{fam}**")
            st.metric("Diferencia GASTO ANUAL", f"{chk['diff_gasto']:,.2f}")
            st.metric("Diferencia MONTO ACTIVADO", f"{chk['diff_act']:,.2f}")
            if abs(chk["diff_gasto"]) > 1 or abs(chk["diff_act"]) > 1:
                all_ok = False
                st.warning("Diferencia > $1")
            else:
                st.info("Cuadratura OK")

    if not all_ok:
        st.warning("Alguna familia presenta una diferencia de cuadratura mayor a $1. Revisa los archivos fuente.")

    st.subheader("Detalle REP_2")
    st.dataframe(
        df.style.format({
            "% NO ACTIVADO ASIGNADO FAMILIA SERVICIOS": "{:.2%}",
            "% ACTIVADO ASIGNADO FAMILIA SERVICIOS": "{:.2%}",
            "GASTO ANUAL": "{:,.0f}",
            "MONTO ACTIVADO": "{:,.0f}",
        }),
        use_container_width=True,
    )

    excel_bytes = build_excel(final_rows, familia_map, f_template.getvalue() if f_template else None)
    st.download_button(
        "Descargar REP_2.xlsx",
        data=excel_bytes,
        file_name="REP_2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Sube los 6 archivos requeridos en el panel izquierdo y presiona **Generar REP_2**.")
