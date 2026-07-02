"""
App Streamlit - Generador REP_2 (Costos y Gastos por Familia de Servicios)

Consolida 5 familias de gasto:
  - GRH: Gastos de Recursos Humanos            (GRH_8 + GRH_11)
  - GCP: Gastos Generales de Personal          (GCP_4 + GCP_5, ya abierta por recurso)
  - GGV: Gastos Generales Vehículos y Equipos  (GGV_4 + GGV_5)
  - GGI: Gastos Generales Bienes Inmuebles     (GGI_5, ya abierta por recurso)
  - GGM: Gastos Generales Bienes Muebles       (GGM_1..GGM_5, sin apertura;
          por defecto 100% -> servicio 1101, con parametrización opcional
          para destinar % a servicios no regulados)

Sube los archivos correspondientes y la app arma la tabla REP_2 consolidada,
lista para descargar, con validaciones de cuadratura por cada familia.
"""

import io
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

SERVICIOS_NO_REGULADOS = {
    2101: "INSPECCIÓN DE INSTALACIÓN DOMICILIARIA",
    2102: "VERIFICACIÓN DE FUNCIONAMIENTO DEL MEDIDOR A TRAVÉS DEL MEDIDOR PATRÓN",
    2103: "REPARACIÓN DE INSTALACIÓN DOMICILIARIA DE AGUA POTABLE",
    2104: "DETECCIÓN DE FUGA INTRADOMICILIARIA DE AGUA POTABLE",
    2105: "CONSTRUCCIÓN DE ARRANQUE",
    2106: "ELIMINACIÓN DE ARRANQUE DE AGUA POTABLE A SOLICITUD DEL CLIENTE",
    2107: "CAMBIO DE MEDIDOR",
    2108: "VENTA DE MEDIDOR",
    2109: "CAMBIO DE UBICACIÓN DE MEDIDOR DE AGUA POTABLE",
    2110: "CONEXIÓN DE REDES A EJECUTAR CON PERSONAL DE LA EMPRESA",
    2111: "DESOBSTRUCCIÓN DE UNIÓN DOMICILIARIA",
    2112: "LIMPIEZA DE FOSA SÉPTICA",
    2113: "CONSTRUCCIÓN DE UD",
    2114: "REPARACIÓN DE CÁMARAS DE INSPECCIÓN DOMICILIARIAS",
    2115: "CONSTRUCCIÓN DE CÁMARAS DE INSPECCIÓN DOMICILIARIAS",
    2116: "EMPALMES DE COLECTORES LOTEOS NUEVOS REALIZADOS POR PERSONAL DE LA EMPRESA",
    2117: "OTRAS PRESTACIONES EXPRESAMENTE VINCULADAS AL INCISO 2° DEL ARTÍCULO 21 DE LA LEY DE TARIFAS",
    2201: "VENTA DE AGUA Y SERVICIOS DE ALCANTARILLADO",
    2202: "VENTA DE SUBPRODUCTOS",
    2203: "TRATAMIENTO DE RILES",
    2204: "SERVICIOS DE ASESORÍA Y GESTIÓN NO APR (PARA TERCEROS, RELACIONADOS O NO)",
    2205: "SERVICIOS DE ASESORÍA Y GESTIÓN DE APR",
    2206: "SERVICIOS DE INGENIERÍA",
    2207: "SERVICIOS DE CONSTRUCCIÓN E INSTALACIÓN",
    2208: "SERVICIOS DE ANÁLISIS DE LABORATORIO",
    2209: "SERVICIOS DE REPARACIONES, OPERACIONES Y MANTENIMIENTOS",
    2210: "SERVICIOS DE INSPECCIONES Y CERTIFICACIONES",
    2211: "VENTA DE EQUIPOS Y PIEZAS",
    2212: "ARRIENDO DE ACTIVOS",
    2213: "MONITOREOS AMBIENTALES",
    2214: "OTROS SERVICIOS NO REGULADOS",
}

RECURSOS_GGM = [2401, 2402, 2403, 2404, 2405, 2406, 2407, 2408, 2409, 2410, 2411]


def read_rows(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    return [r for r in ws.iter_rows(min_row=2, values_only=True) if r[4] is not None]


def familia(cod_servicio):
    return cod_servicio // 100


def build_rep2(files_bytes, ggm_params):
    """
    files_bytes: dict con claves grh8, grh11, gcp4, gcp5, ggv4, ggv5, ggi5,
                 ggm1..ggm5 -> bytes del archivo
    ggm_params: dict {cod_recurso: [(cod_servicio_no_regulado, pct), ...]}
    """
    grh8 = read_rows(files_bytes["grh8"])
    grh11 = read_rows(files_bytes["grh11"])
    gcp4 = read_rows(files_bytes["gcp4"])
    gcp5 = read_rows(files_bytes["gcp5"])
    ggv4 = read_rows(files_bytes["ggv4"])
    ggv5 = read_rows(files_bytes["ggv5"])
    ggi5 = read_rows(files_bytes["ggi5"])
    ggm_tablas = [read_rows(files_bytes[f"ggm{i}"]) for i in range(1, 6)]

    agg = defaultdict(lambda: [0.0, 0.0])
    EMPRESA, PERIODO, ANIO, SECTOR = grh8[0][0], grh8[0][1], grh8[0][2], grh8[0][3]

    # --- GRH ---
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

    # --- GCP ---
    for r in gcp5:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        fam = familia(cod_serv)
        k = (empresa, periodo, anio, sector, cod_recurso, fam)
        agg[k][0] += gasto_no_act

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

    # --- GGV ---
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

    # --- GGI ---
    for r in ggi5:
        empresa, periodo, anio, sector, id_inmueble, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r[:10]
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        fam = familia(cod_serv)
        k = (empresa, periodo, anio, sector, cod_recurso, fam)
        agg[k][0] += gasto_no_act

    # --- GGM ---
    ggm_by_recurso = defaultdict(lambda: [0.0, 0.0])
    for tabla in ggm_tablas:
        for r in tabla:
            cod_recurso = r[4]
            monto_act = r[-3]
            total_gasto = r[-1]
            ggm_by_recurso[cod_recurso][0] += total_gasto
            ggm_by_recurso[cod_recurso][1] += monto_act

    for cod_recurso, (gasto_no_act, monto_act) in ggm_by_recurso.items():
        overrides = ggm_params.get(cod_recurso, [])
        pct_reg = 1.0 - sum(p for _, p in overrides)
        fam = familia(1101)
        k = (EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam)
        agg[k][0] += gasto_no_act * pct_reg
        agg[k][1] += monto_act * pct_reg
        for cod_serv_noreg, pct in overrides:
            fam2 = familia(cod_serv_noreg)
            k2 = (EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam2)
            agg[k2][0] += gasto_no_act * pct
            agg[k2][1] += monto_act * pct

    # --- Normalizar % por recurso ---
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

    # --- Validaciones por familia ---
    recursos_grh = set(r[6] for r in grh8)
    recursos_gcp = set(r[6] for r in gcp4)
    recursos_ggv = set(r[5] for r in ggv4)
    recursos_ggi = set(r[5] for r in ggi5)
    recursos_ggm = set(ggm_by_recurso.keys())

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
        "GGI": validar(recursos_ggi, sum(r[6] for r in ggi5), 0),
        "GGM": validar(recursos_ggm, sum(v[0] for v in ggm_by_recurso.values()), sum(v[1] for v in ggm_by_recurso.values())),
    }

    familia_map = {}
    for c in recursos_grh:
        familia_map[c] = "GRH - Gastos Recursos Humanos"
    for c in recursos_gcp:
        familia_map[c] = "GCP - Gastos Generales de Personal"
    for c in recursos_ggv:
        familia_map[c] = "GGV - Gastos Generales Vehículos y Equipos"
    for c in recursos_ggi:
        familia_map[c] = "GGI - Gastos Generales Bienes Inmuebles"
    for c in recursos_ggm:
        familia_map[c] = "GGM - Gastos Generales Bienes Muebles"

    return final_rows, checks, familia_map, ggm_by_recurso, recursos_ggm


def build_excel(final_rows, familia_map, ggm_by_recurso, recursos_ggm, ggm_params, template_bytes=None):
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

    # Resumen por familia de gasto
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

    orden_familias = [
        "GRH - Gastos Recursos Humanos", "GCP - Gastos Generales de Personal",
        "GGV - Gastos Generales Vehículos y Equipos", "GGI - Gastos Generales Bienes Inmuebles",
        "GGM - Gastos Generales Bienes Muebles",
    ]
    for fg in orden_familias:
        if fg in resumen:
            g, a = resumen[fg]
            ws2.append([fg, round(g, 2), round(a, 2), round(g + a, 2)])
    total_g = sum(v[0] for v in resumen.values())
    total_a = sum(v[1] for v in resumen.values())
    ws2.append(["TOTAL REP_2", round(total_g, 2), round(total_a, 2), round(total_g + total_a, 2)])

    for r in range(2, ws2.max_row + 1):
        for c in range(1, 5):
            cell = ws2.cell(row=r, column=c)
            cell.font = Font(name="Arial", size=10, bold=(r == ws2.max_row))
            cell.border = border
            if c in (2, 3, 4):
                cell.number_format = '#,##0;(#,##0);"-"'
    ws2.column_dimensions["A"].width = 42
    for col in ["B", "C", "D"]:
        ws2.column_dimensions[col].width = 20

    # Parametrización GGM aplicada
    ws3 = wb.create_sheet("Parametrizacion_GGM")
    ws3.append(["CÓDIGO RECURSO GGM", "GASTO ANUAL TOTAL (100%)", "CÓDIGO SERVICIO NO REGULADO", "SERVICIO NO REGULADO", "% ASIGNADO"])
    for c in range(1, 6):
        cell = ws3.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws3.row_dimensions[1].height = 30

    for cod_recurso in sorted(recursos_ggm):
        gasto_total = ggm_by_recurso[cod_recurso][0] + ggm_by_recurso[cod_recurso][1]
        overrides = ggm_params.get(cod_recurso, [])
        if not overrides:
            ws3.append([cod_recurso, round(gasto_total, 2), "(sin parametrizar -> 100% a servicio 1101)", "", ""])
        else:
            for cod_serv, pct in overrides:
                ws3.append([cod_recurso, round(gasto_total, 2), cod_serv, SERVICIOS_NO_REGULADOS.get(cod_serv, ""), pct])

    for r in range(2, ws3.max_row + 1):
        for c in range(1, 6):
            cell = ws3.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
            if c == 2:
                cell.number_format = '#,##0;(#,##0);"-"'
            if c == 5 and isinstance(ws3.cell(row=r, column=5).value, float):
                cell.number_format = "0.00%"
    ws3.column_dimensions["A"].width = 18
    ws3.column_dimensions["B"].width = 22
    ws3.column_dimensions["C"].width = 42
    ws3.column_dimensions["D"].width = 55
    ws3.column_dimensions["E"].width = 14

    # Catálogo de servicios no regulados
    ws4 = wb.create_sheet("Catalogo_Servicios_No_Regulados")
    ws4.append(["CÓDIGO SERVICIO NO REGULADO", "SERVICIO NO REGULADO"])
    for c in range(1, 3):
        cell = ws4.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    for cod, desc in SERVICIOS_NO_REGULADOS.items():
        ws4.append([cod, desc])
    for r in range(2, ws4.max_row + 1):
        for c in range(1, 3):
            cell = ws4.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
    ws4.column_dimensions["A"].width = 22
    ws4.column_dimensions["B"].width = 70

    # Diccionario SISS (si se subió plantilla)
    if template_bytes is not None:
        try:
            tmpl = openpyxl.load_workbook(io.BytesIO(template_bytes), data_only=True)
            src_name = "Diccionario_REP_2" if "Diccionario_REP_2" in tmpl.sheetnames else tmpl.sheetnames[0]
            src_ws = tmpl[src_name]
            dst_ws = wb.create_sheet("Diccionario_REP_2")
            for row in src_ws.iter_rows():
                for cell in row:
                    dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
        except Exception:
            pass

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ---------------------------------------------------------------- UI --------
st.title("Generador REP_2 · Costos y Gastos por Familia de Servicios")
st.caption(
    "Consolida las familias de gasto GRH, GCP, GGV, GGI y GGM en la tabla "
    "REP_2 exigida por la SISS."
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
    st.markdown("**GGI — Bienes Inmuebles**")
    f_ggi5 = st.file_uploader("GGI_5.xlsx", type="xlsx", key="ggi5")
    st.markdown("**GGM — Bienes Muebles**")
    f_ggm1 = st.file_uploader("GGM_1.xlsx", type="xlsx", key="ggm1")
    f_ggm2 = st.file_uploader("GGM_2.xlsx", type="xlsx", key="ggm2")
    f_ggm3 = st.file_uploader("GGM_3.xlsx", type="xlsx", key="ggm3")
    f_ggm4 = st.file_uploader("GGM_4.xlsx", type="xlsx", key="ggm4")
    f_ggm5 = st.file_uploader("GGM_5.xlsx", type="xlsx", key="ggm5")
    st.markdown("**Plantilla (opcional)**")
    f_template = st.file_uploader("REP_2.xlsx (diccionario)", type="xlsx", key="template")

    required = [f_grh8, f_grh11, f_gcp4, f_gcp5, f_ggv4, f_ggv5, f_ggi5,
                f_ggm1, f_ggm2, f_ggm3, f_ggm4, f_ggm5]
    all_ready = all(required)

st.markdown(
    """
**Lógica aplicada**
- **GRH**: % de dedicación por persona (GRH_11) aplicado a GRH_8.
- **GCP**: GCP_5 ya viene abierta por recurso y servicio — gasto no activado
  directo; el activado usa el mismo % aplicado a GCP_4.
- **GGV**: % de dedicación por activo (GGV_5) aplicado a GGV_4.
- **GGI**: GGI_5 ya viene abierta por recurso y servicio, igual que GCP_5.
- **GGM**: 5 tablas planas (GGM_1 a GGM_5, recursos 2401-2411) sin apertura
  por servicio. Por defecto, 100% va al servicio regulado 1101. Puedes
  parametrizar abajo qué % de cada recurso GGM se destina a un servicio no
  regulado; el remanente sigue yendo a 1101.
    """
)

st.subheader("Parametrización opcional — GGM")
st.caption(
    "Por defecto todo el gasto GGM se asigna 100% al servicio regulado 1101. "
    "Si quieres destinar un % de algún recurso GGM (2401 a 2411) a un servicio "
    "no regulado, agrégalo aquí. El resto del recurso sigue yendo a 1101."
)

if "ggm_overrides" not in st.session_state:
    st.session_state.ggm_overrides = []

with st.expander("Configurar parametrización GGM", expanded=False):
    col1, col2, col3, col4 = st.columns([1.2, 2.5, 1, 0.8])
    with col1:
        sel_recurso = st.selectbox("Código Recurso GGM", RECURSOS_GGM, key="sel_recurso_ggm")
    with col2:
        sel_servicio = st.selectbox(
            "Servicio no regulado destino",
            options=list(SERVICIOS_NO_REGULADOS.keys()),
            format_func=lambda c: f"{c} - {SERVICIOS_NO_REGULADOS[c]}",
            key="sel_servicio_ggm",
        )
    with col3:
        sel_pct = st.number_input("% asignado", min_value=0.0, max_value=100.0, value=10.0, step=1.0, key="sel_pct_ggm")
    with col4:
        st.write("")
        st.write("")
        if st.button("Agregar"):
            st.session_state.ggm_overrides.append((sel_recurso, sel_servicio, sel_pct / 100.0))

    if st.session_state.ggm_overrides:
        st.markdown("**Parametrización actual:**")
        for i, (cod_r, cod_s, pct) in enumerate(st.session_state.ggm_overrides):
            c1, c2, c3, c4 = st.columns([1.2, 2.5, 1, 0.8])
            c1.write(cod_r)
            c2.write(f"{cod_s} - {SERVICIOS_NO_REGULADOS.get(cod_s, '')}")
            c3.write(f"{pct:.1%}")
            if c4.button("Quitar", key=f"del_{i}"):
                st.session_state.ggm_overrides.pop(i)
                st.rerun()
    else:
        st.info("Sin parametrización: 100% de todos los recursos GGM va al servicio 1101.")

ggm_params = defaultdict(list)
for cod_r, cod_s, pct in st.session_state.ggm_overrides:
    ggm_params[cod_r].append((cod_s, pct))

# Validar que ningún recurso supere 100% en overrides
overflow = {r: sum(p for _, p in lst) for r, lst in ggm_params.items() if sum(p for _, p in lst) > 1.0}
if overflow:
    st.error(f"La suma de % parametrizados supera 100% para el(los) recurso(s): {list(overflow.keys())}. Ajusta los valores.")

run = st.button("Generar REP_2", type="primary", disabled=not all_ready or bool(overflow))

if not all_ready:
    st.info("Sube los 12 archivos requeridos en el panel izquierdo (GGV_5 y GGI_5 incluidos) para habilitar la generación.")

if run:
    files_bytes = {
        "grh8": f_grh8.getvalue(), "grh11": f_grh11.getvalue(),
        "gcp4": f_gcp4.getvalue(), "gcp5": f_gcp5.getvalue(),
        "ggv4": f_ggv4.getvalue(), "ggv5": f_ggv5.getvalue(),
        "ggi5": f_ggi5.getvalue(),
        "ggm1": f_ggm1.getvalue(), "ggm2": f_ggm2.getvalue(), "ggm3": f_ggm3.getvalue(),
        "ggm4": f_ggm4.getvalue(), "ggm5": f_ggm5.getvalue(),
    }
    try:
        final_rows, checks, familia_map, ggm_by_recurso, recursos_ggm = build_rep2(files_bytes, dict(ggm_params))
    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
        st.stop()

    df = pd.DataFrame(final_rows, columns=HEADERS)
    st.success(f"REP_2 generado con {len(df)} filas (5 familias de gasto consolidadas).")

    st.subheader("Validación de cuadratura por familia de gasto")
    cols = st.columns(5)
    all_ok = True
    for col, (fam, chk) in zip(cols, checks.items()):
        with col:
            st.markdown(f"**{fam}**")
            st.metric("Δ GASTO ANUAL", f"{chk['diff_gasto']:,.2f}")
            st.metric("Δ MONTO ACTIVADO", f"{chk['diff_act']:,.2f}")
            if abs(chk["diff_gasto"]) > 1 or abs(chk["diff_act"]) > 1:
                all_ok = False
                st.warning("Diferencia > $1")
            else:
                st.info("OK")

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

    excel_bytes = build_excel(
        final_rows, familia_map, ggm_by_recurso, recursos_ggm, dict(ggm_params),
        f_template.getvalue() if f_template else None
    )
    st.download_button(
        "Descargar REP_2.xlsx",
        data=excel_bytes,
        file_name="REP_2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
