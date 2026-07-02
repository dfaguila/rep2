"""
App Streamlit - Generador REP_2 (Costos y Gastos por Familia de Servicios)

Consolida 7 familias de gasto:
  - GRH: Recursos Humanos                (GRH_8 + GRH_11)
  - GCP: Gastos Generales de Personal    (GCP_4 + GCP_5, ya abierta por recurso)
  - GGV: Vehículos y Equipos             (GGV_4 + GGV_5)
  - GGI: Bienes Inmuebles                (GGI_5, ya abierta por recurso y servicio)
  - GGM: Bienes Muebles                  (GGM_1..GGM_5, sin apertura, recursos 2401-2411)
  - OGG: Otros Gastos Generales          (OGG_5, sin apertura, recursos 2501-2550)
  - MEI: Materiales e Insumos            (MEI_1..MEI_4, sin apertura, recursos 4101-4106)

GGM, OGG y MEI no traen columna de servicio: por defecto 100% se asigna al
servicio regulado 1101, con parametrización opcional para destinar % a
servicios no regulados.
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

# Recursos parametrizables por familia plana (para el selector de la UI)
RECURSOS_GGM = list(range(2401, 2412))
RECURSOS_OGG = list(range(2501, 2551))
RECURSOS_MEI = [4101, 4102, 4103, 4104, 4105, 4106]


def read_rows(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    return [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]


def familia(cod_servicio):
    return cod_servicio // 100


def procesar_familia_plana(agg, EMPRESA, PERIODO, ANIO, SECTOR, tablas_specs, params):
    by_recurso = defaultdict(lambda: [0.0, 0.0])
    for filas, idx_r, idx_a, idx_g in tablas_specs:
        for r in filas:
            cod_recurso = r[idx_r]
            monto_act = r[idx_a]
            total_gasto = r[idx_g]
            by_recurso[cod_recurso][0] += total_gasto
            by_recurso[cod_recurso][1] += monto_act

    for cod_recurso, (gasto_no_act, monto_act) in by_recurso.items():
        overrides = params.get(cod_recurso, [])
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
    return by_recurso


def build_rep2(fb, ggm_params, ogg_params, mei_params):
    grh8 = read_rows(fb["grh8"])
    grh11 = read_rows(fb["grh11"])
    gcp4 = read_rows(fb["gcp4"])
    gcp5 = read_rows(fb["gcp5"])
    ggv4 = read_rows(fb["ggv4"])
    ggv5 = read_rows(fb["ggv5"])
    ggi5 = read_rows(fb["ggi5"])
    ggm_tablas = [read_rows(fb[f"ggm{i}"]) for i in range(1, 6)]
    ogg5 = read_rows(fb["ogg5"])
    mei1 = read_rows(fb["mei1"])
    mei2 = read_rows(fb["mei2"])
    mei3 = read_rows(fb["mei3"])
    mei4 = read_rows(fb["mei4"])

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
    ggm_by_recurso = procesar_familia_plana(
        agg, EMPRESA, PERIODO, ANIO, SECTOR,
        [(t, 4, -3, -1) for t in ggm_tablas],
        ggm_params
    )

    # --- OGG ---
    ogg_by_recurso = procesar_familia_plana(
        agg, EMPRESA, PERIODO, ANIO, SECTOR,
        [(ogg5, 4, 9, 11)],
        ogg_params
    )

    # --- MEI (índices de columna distintos por tabla; MEI_1 recurso en col6) ---
    mei_by_recurso = procesar_familia_plana(
        agg, EMPRESA, PERIODO, ANIO, SECTOR,
        [(mei1, 6, 20, 22), (mei2, 4, 17, 16), (mei3, 4, 17, 19), (mei4, 4, 14, 16)],
        mei_params
    )

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

    # --- Validaciones ---
    recursos_grh = set(r[6] for r in grh8)
    recursos_gcp = set(r[6] for r in gcp4)
    recursos_ggv = set(r[5] for r in ggv4)
    recursos_ggi = set(r[5] for r in ggi5)
    recursos_ggm = set(ggm_by_recurso.keys())
    recursos_ogg = set(ogg_by_recurso.keys())
    recursos_mei = set(mei_by_recurso.keys())

    def validar(recursos, sum_gasto_fuente, sum_act_fuente):
        sum_gasto_rep2 = sum(r[7] for r in final_rows if r[4] in recursos)
        sum_act_rep2 = sum(r[9] for r in final_rows if r[4] in recursos)
        return {
            "diff_gasto": sum_gasto_rep2 - sum_gasto_fuente,
            "diff_act": sum_act_rep2 - sum_act_fuente,
        }

    checks = {
        "GRH": validar(recursos_grh, sum(r[10] for r in grh8), sum(r[8] for r in grh8)),
        "GCP": validar(recursos_gcp, sum(r[10] for r in gcp4), sum(r[8] for r in gcp4)),
        "GGV": validar(recursos_ggv, sum(r[9] for r in ggv4), sum(r[7] for r in ggv4)),
        "GGI": validar(recursos_ggi, sum(r[6] for r in ggi5), 0),
        "GGM": validar(recursos_ggm, sum(v[0] for v in ggm_by_recurso.values()), sum(v[1] for v in ggm_by_recurso.values())),
        "OGG": validar(recursos_ogg, sum(v[0] for v in ogg_by_recurso.values()), sum(v[1] for v in ogg_by_recurso.values())),
        "MEI": validar(recursos_mei, sum(v[0] for v in mei_by_recurso.values()), sum(v[1] for v in mei_by_recurso.values())),
    }

    familia_map = {}
    for c in recursos_grh: familia_map[c] = "GRH - Gastos Recursos Humanos"
    for c in recursos_gcp: familia_map[c] = "GCP - Gastos Generales de Personal"
    for c in recursos_ggv: familia_map[c] = "GGV - Gastos Generales Vehículos y Equipos"
    for c in recursos_ggi: familia_map[c] = "GGI - Gastos Generales Bienes Inmuebles"
    for c in recursos_ggm: familia_map[c] = "GGM - Gastos Generales Bienes Muebles"
    for c in recursos_ogg: familia_map[c] = "OGG - Otros Gastos Generales"
    for c in recursos_mei: familia_map[c] = "MEI - Materiales e Insumos"

    by_recurso_planas = {"GGM": ggm_by_recurso, "OGG": ogg_by_recurso, "MEI": mei_by_recurso}
    return final_rows, checks, familia_map, by_recurso_planas


def build_excel(final_rows, familia_map, by_recurso_planas, params_by_familia, template_bytes=None):
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

    # Resumen por familia
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
        "GGM - Gastos Generales Bienes Muebles", "OGG - Otros Gastos Generales",
        "MEI - Materiales e Insumos",
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

    # Parametrización unificada GGM+OGG+MEI
    ws3 = wb.create_sheet("Parametrizacion_familias_planas")
    ws3.append(["Familia", "CÓDIGO RECURSO", "GASTO ANUAL TOTAL (100%)", "CÓDIGO SERVICIO NO REGULADO", "SERVICIO NO REGULADO", "% ASIGNADO"])
    for c in range(1, 7):
        cell = ws3.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws3.row_dimensions[1].height = 30

    for nombre_familia, by_recurso in by_recurso_planas.items():
        params = params_by_familia.get(nombre_familia, {})
        for cod_recurso in sorted(by_recurso.keys()):
            gasto_total = by_recurso[cod_recurso][0] + by_recurso[cod_recurso][1]
            overrides = params.get(cod_recurso, [])
            if not overrides:
                ws3.append([nombre_familia, cod_recurso, round(gasto_total, 2), "(sin parametrizar -> 100% a servicio 1101)", "", ""])
            else:
                for cod_serv, pct in overrides:
                    ws3.append([nombre_familia, cod_recurso, round(gasto_total, 2), cod_serv, SERVICIOS_NO_REGULADOS.get(cod_serv, ""), pct])

    for r in range(2, ws3.max_row + 1):
        for c in range(1, 7):
            cell = ws3.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
            if c == 3:
                cell.number_format = '#,##0;(#,##0);"-"'
            if c == 6 and isinstance(ws3.cell(row=r, column=6).value, float):
                cell.number_format = "0.00%"
    ws3.column_dimensions["A"].width = 10
    ws3.column_dimensions["B"].width = 16
    ws3.column_dimensions["C"].width = 22
    ws3.column_dimensions["D"].width = 42
    ws3.column_dimensions["E"].width = 55
    ws3.column_dimensions["F"].width = 12

    # Catálogo
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

    # Diccionario SISS
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


def panel_parametrizacion(nombre, recursos_disponibles, session_key):
    """Renderiza un panel de parametrización reutilizable para GGM/OGG/MEI."""
    if session_key not in st.session_state:
        st.session_state[session_key] = []

    st.markdown(f"**{nombre}**")
    col1, col2, col3, col4 = st.columns([1.2, 2.5, 1, 0.8])
    with col1:
        sel_recurso = st.selectbox("Código Recurso", recursos_disponibles, key=f"sel_recurso_{session_key}")
    with col2:
        sel_servicio = st.selectbox(
            "Servicio no regulado destino",
            options=list(SERVICIOS_NO_REGULADOS.keys()),
            format_func=lambda c: f"{c} - {SERVICIOS_NO_REGULADOS[c]}",
            key=f"sel_servicio_{session_key}",
        )
    with col3:
        sel_pct = st.number_input("% asignado", min_value=0.0, max_value=100.0, value=10.0, step=1.0, key=f"sel_pct_{session_key}")
    with col4:
        st.write("")
        st.write("")
        if st.button("Agregar", key=f"add_{session_key}"):
            st.session_state[session_key].append((sel_recurso, sel_servicio, sel_pct / 100.0))

    if st.session_state[session_key]:
        for i, (cod_r, cod_s, pct) in enumerate(st.session_state[session_key]):
            c1, c2, c3, c4 = st.columns([1.2, 2.5, 1, 0.8])
            c1.write(cod_r)
            c2.write(f"{cod_s} - {SERVICIOS_NO_REGULADOS.get(cod_s, '')}")
            c3.write(f"{pct:.1%}")
            if c4.button("Quitar", key=f"del_{session_key}_{i}"):
                st.session_state[session_key].pop(i)
                st.rerun()
    else:
        st.caption(f"Sin parametrización: 100% de todos los recursos {nombre} va al servicio 1101.")

    params = defaultdict(list)
    for cod_r, cod_s, pct in st.session_state[session_key]:
        params[cod_r].append((cod_s, pct))
    return dict(params)


# ---------------------------------------------------------------- UI --------
st.title("Generador REP_2 · Costos y Gastos por Familia de Servicios")
st.caption(
    "Consolida las familias GRH, GCP, GGV, GGI, GGM, OGG y MEI en la tabla "
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
    st.markdown("**OGG — Otros Gastos Generales**")
    f_ogg5 = st.file_uploader("OGG_5.xlsx", type="xlsx", key="ogg5")
    st.markdown("**MEI — Materiales e Insumos**")
    f_mei1 = st.file_uploader("MEI_1.xlsx", type="xlsx", key="mei1")
    f_mei2 = st.file_uploader("MEI_2.xlsx", type="xlsx", key="mei2")
    f_mei3 = st.file_uploader("MEI_3.xlsx", type="xlsx", key="mei3")
    f_mei4 = st.file_uploader("MEI_4.xlsx", type="xlsx", key="mei4")
    st.markdown("**Plantilla (opcional)**")
    f_template = st.file_uploader("REP_2.xlsx (diccionario)", type="xlsx", key="template")

    required = [f_grh8, f_grh11, f_gcp4, f_gcp5, f_ggv4, f_ggv5, f_ggi5,
                f_ggm1, f_ggm2, f_ggm3, f_ggm4, f_ggm5, f_ogg5,
                f_mei1, f_mei2, f_mei3, f_mei4]
    all_ready = all(required)

st.markdown(
    """
**Lógica aplicada**
- **GRH / GGV**: % de dedicación (por persona o por activo) aplicado a los montos de gasto.
- **GCP / GGI**: ya vienen abiertas por recurso y servicio.
- **GGM / OGG / MEI**: sin apertura por servicio — por defecto 100% va al
  servicio regulado 1101; se puede parametrizar % a servicios no regulados
  en los paneles de abajo.
    """
)

st.subheader("Parametrización opcional — familias sin apertura por servicio")
st.caption("Por defecto todo el gasto se asigna 100% al servicio regulado 1101.")

with st.expander("Configurar parametrización GGM / OGG / MEI", expanded=False):
    ggm_params = panel_parametrizacion("GGM (recursos 2401-2411)", RECURSOS_GGM, "ggm_overrides")
    st.divider()
    ogg_params = panel_parametrizacion("OGG (recursos 2501-2550)", RECURSOS_OGG, "ogg_overrides")
    st.divider()
    mei_params = panel_parametrizacion("MEI (recursos 4101-4106)", RECURSOS_MEI, "mei_overrides")

if "ggm_overrides" not in st.session_state:
    ggm_params = {}
if "ogg_overrides" not in st.session_state:
    ogg_params = {}
if "mei_overrides" not in st.session_state:
    mei_params = {}

# Validar overflow > 100%
def check_overflow(params):
    return {r: sum(p for _, p in lst) for r, lst in params.items() if sum(p for _, p in lst) > 1.0}

overflow = {}
overflow.update(check_overflow(ggm_params))
overflow.update(check_overflow(ogg_params))
overflow.update(check_overflow(mei_params))
if overflow:
    st.error(f"La suma de % parametrizados supera 100% para el(los) recurso(s): {list(overflow.keys())}. Ajusta los valores.")

run = st.button("Generar REP_2", type="primary", disabled=not all_ready or bool(overflow))

if not all_ready:
    st.info("Sube los 17 archivos requeridos en el panel izquierdo para habilitar la generación.")

if run:
    fb = {
        "grh8": f_grh8.getvalue(), "grh11": f_grh11.getvalue(),
        "gcp4": f_gcp4.getvalue(), "gcp5": f_gcp5.getvalue(),
        "ggv4": f_ggv4.getvalue(), "ggv5": f_ggv5.getvalue(),
        "ggi5": f_ggi5.getvalue(),
        "ggm1": f_ggm1.getvalue(), "ggm2": f_ggm2.getvalue(), "ggm3": f_ggm3.getvalue(),
        "ggm4": f_ggm4.getvalue(), "ggm5": f_ggm5.getvalue(),
        "ogg5": f_ogg5.getvalue(),
        "mei1": f_mei1.getvalue(), "mei2": f_mei2.getvalue(),
        "mei3": f_mei3.getvalue(), "mei4": f_mei4.getvalue(),
    }
    try:
        final_rows, checks, familia_map, by_recurso_planas = build_rep2(fb, ggm_params, ogg_params, mei_params)
    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
        st.stop()

    df = pd.DataFrame(final_rows, columns=HEADERS)
    st.success(f"REP_2 generado con {len(df)} filas (7 familias de gasto consolidadas).")

    st.subheader("Validación de cuadratura por familia de gasto")
    cols = st.columns(7)
    all_ok = True
    for col, (fam, chk) in zip(cols, checks.items()):
        with col:
            st.markdown(f"**{fam}**")
            st.metric("Δ GASTO", f"{chk['diff_gasto']:,.1f}")
            st.metric("Δ ACTIVADO", f"{chk['diff_act']:,.1f}")
            if abs(chk["diff_gasto"]) > 1 or abs(chk["diff_act"]) > 1:
                all_ok = False
                st.warning("Diff > $1")
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

    params_by_familia = {"GGM": ggm_params, "OGG": ogg_params, "MEI": mei_params}
    excel_bytes = build_excel(
        final_rows, familia_map, by_recurso_planas, params_by_familia,
        f_template.getvalue() if f_template else None
    )
    st.download_button(
        "Descargar REP_2.xlsx",
        data=excel_bytes,
        file_name="REP_2.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
