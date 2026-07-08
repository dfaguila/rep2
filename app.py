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
import re
import difflib
import statistics
from collections import defaultdict

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule

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

# ============================================================================
# Diccionario de códigos de recurso (para el Panel de Visualización)
# ============================================================================
RECURSO_NOMBRE = {
    1101: "Remuneraciones",
    1102: "Honorarios",
    1103: "Horas Extras",
    1201: "Indemnizaciones",
    1202: "Seguro de Cesantía",
    1203: "Seguro de Accidentes",
    1204: "Seguro de Invalidez y Sobrevivencia",
    1205: "Otros Beneficios Adicionales",
    2101: "Alimentacíón",
    2102: "Capacitación",
    2103: "Pasajes",
    2104: "Alojamientos",
    2105: "Viáticos",
    2106: "Accesorios de personal",
    2201: "Arriendo de vehículos y maquinarias",
    2202: "Combustible",
    2203: "Permisos de circulación",
    2204: "Revisión técnica",
    2205: "Seguros",
    2206: "Mantención Preventiva (no incluye combustible)",
    2207: "Recobros y Mantención Correctiva",
    2208: "Peajes",
    2209: "Tags",
    2210: "Implementación tag",
    2211: "Gasto identificación de vehículos",
    2301: "Arriendo de inmuebles",
    2302: "Consumos básicos",
    2303: "Servicio de aseo",
    2304: "Materiales de aseo",
    2305: "Mantención de inmuebles",
    2306: "Mantención de extintores",
    2307: "Mantención de areas verdes",
    2308: "Vigilancia presencial",
    2309: "Vigilancia a distancia",
    2310: "Contribuciones",
    2311: "Comisiones de Corretaje de Inmuebles por compra o arriendo",
    2401: "Arriendo de equipos informáticos",
    2402: "Servicios informáticos",
    2403: "Telefonía Fija",
    2404: "Enlasces de Internet Fijos",
    2405: "Enlaces de datos entre inmuebles (incluye enlaces satelitales)",
    2406: "Telefonía y Banda Ancha Móvil",
    2407: "Tarjetas SIM para aplicaciones M2M",
    2408: "Radio trunking de voz y datos",
    2409: "Telefonía Satelital",
    2410: "Materiales e insumos de oficina, computacionales y bodega",
    2411: "Materiales e insumos de laboratorio",
    2501: "Dietas del Directorio",
    2502: "Gastos de Representación Directorio",
    2503: "Patentes Comerciales",
    2504: "Servicios de Imprenta, Fotocopiado y Reproducción",
    2505: "Trámites y gastos notariales",
    2506: "Actuaciones Judiciales",
    2507: "Inscripciones",
    2508: "Peritajes",
    2509: "Tasaciones",
    2510: "Enlaces Satelitales",
    2511: "Líneas Transmisión de Datos",
    2512: "Enlaces de Internet",
    2513: "Impuestos por uso de Canales de Radiofrecuencia",
    2514: "Fletes",
    2515: "Transporte de Correspondencia (incluye servicios postales y mensajería)",
    2516: "Transporte de Muestras de Laboratorio",
    2517: "Seguros de infraestructura de capacidad",
    2518: "Seguros de inmuebles",
    2519: "Seguros de redes",
    2520: "Autoseguro",
    2521: "Seguros de Responsabilidad Civil",
    2522: "Seguros Menores",
    2523: "Deducibles pagados",
    2524: "Publicidad y Avisos (Radio, TV, Diarios u otros medios)",
    2525: "Diseño gráfico",
    2526: "Rotulaciones gráficas",
    2527: "Materiales de difusión",
    2528: "Auspicios y aportes",
    2529: "Donaciones",
    2530: "Eventos comunitarios",
    2531: "Eventos corporativos (juntas de accionistas, cenas fin de año, etc.)",
    2532: "Campañas de educación",
    2533: "Materiales de campañas",
    2534: "Derechos de asociaciones y canalistas",
    2535: "Derechos de afiliaciones",
    2536: "Derechos SERVIU o Vialidad",
    2537: "Permisos municipales",
    2538: "Canon anual por activos en comodato",
    2539: "Fondo Fijo Rotativo",
    2540: "Garantías a favor de SISS",
    2541: "Gastos financieros asociados a Garantías SERVIU o Vialidad",
    2542: "Multas",
    2543: "Indemnizaciones a terceros",
    2544: "Suscripciones",
    2545: "Impuestos pagados",
    2546: "Servicios Bancarios",
    2547: "Operaciones Financieras",
    2548: "Castigo incobrables",
    2549: "Otros Gastos Generales",
    3101: "Lectura de medidores",
    3102: "Reparto de boletas y otros documentos",
    3103: "Suministro e impresión de boletas y otros documentos",
    3104: "Servicios de recaudación en cajas externas",
    3105: "Servicios de recaudación en cajas propias",
    3106: "Servicios de atención telefónica o distante",
    3107: "Servicios de inspección comercial",
    3108: "Servicios de cobranza prejudicial",
    3109: "Servicios de gestión",
    3110: "Servicios de transporte de personas (buses de acercamiento, radiotaxis, etc.)",
    3111: "Servicios de almacenamiento y bodegaje",
    3112: "Servicios de traslados de mercancías",
    3113: "Servicios de Procesamiento, Archivo y Digitación de Datos",
    3114: "Auditorías a los Estados Financieros",
    3115: "Clasificación de Riesgo",
    3116: "Administración del Registro de Accionistas",
    3117: "Asesorías Tributarias y Contables",
    3118: "Gestión de Recursos Hídricos",
    3119: "Administración del Rol Privado",
    3120: "Selección de Personal",
    3121: "Auditorías Sistemas de Calidad",
    3122: "Asesorías en Servicio al Cliente",
    3123: "Planes de Desarrollo",
    3124: "Estudios Tarifarios",
    3125: "Comisiones de Expertos en Procesos Tarifarios",
    3126: "Defensa por acciones de responsabilidad civil",
    3127: "Defensa de derechos sobre inmuebles",
    3128: "Defensa en juicios laborales",
    3129: "Reclamaciones tributarias",
    3130: "Asesoría y defensa en procesos penales",
    3131: "Laboral Permanente y Negociación Colectiva",
    3132: "Informes Legales o en Derecho",
    3133: "Otros Servicios No Operacionales",
    4101: "Productos químicos",
    4102: "Energía Eléctrica",
    4103: "Materiales y repuestos",
    4104: "Compra de agua cruda",
    4105: "Compra de agua potable",
    4106: "Arriendo de derechos de agua",
    5101: "Servicios de control de calidad de agua potable",
    5102: "Servicios de control de calidad de agua servidas",
    5103: "Servicios de interconexión AP",
    5104: "Servicios de interconexión AS",
    5105: "Servicios de operación de redes y conexiones",
    5106: "Servicios de operación de infraestructura",
    5107: "Servicios de transporte y disposición de lodos",
    5108: "Servicios de control y monitoreo ambiental",
    5109: "Servicios de mantención de infraestructura",
    5110: "Servicios de mantención de redes y conexiones",
    5111: "Servicios de mantención de recintos",
    5112: "Servicios de mantención de servidumbres",
    5113: "Concesiones marítimas",
    5114: "Otros Servicios Operacionales",
    6101: "Servicios de Terceros asociados a Control Directo de Riles",
    6201: "Servicios de Terceros asociados a Mantención de Grifos",
    6301: "Servicios de Terceros asociados a Corte y Reposición",
    6401: "Servicios de Terceros asociados a Revisión de Proyectos de Construcción",
    6501: "Servicios de Terceros asociados a Verificación de Medidores",
    6601: "Servicios de Terceros asociados a Otras Prestaciones Asociadas",
    7101: "Servicios de Terceros asociados a Estudios preliminares,  Hidrológicos e Hidrogeológicos",
    7102: "Servicios de Terceros asociados a Diseños de obras",
    7103: "Servicios de Terceros asociados a Impacto Ambiental",
    7201: "Servicios de Terceros asociados a Construcción de Obras (Obras Civiles, Eléctricas, de Control, Montaje de Equipos, Mitigación Ambiental, Compensación Ambiental, etc.)",
    7301: "Servicios de Terceros asociados a Inspección Técnica de Obras",
    7401: "Servicios de Terceros asociados a Inversiones en Telemetría y Telecontrol",
    7402: "Servicios de Terceros asociados a Inversiones en Comunicaciones y Sistemas de Información",
    7501: "Servicios de Terceros asociados a Adquisición Inmuebles",
    7502: "Servicios de Terceros asociados a Adquisición Vehículos y equipos",
    7503: "Servicios de Terceros asociados a Adquisición Bienes Muebles e Insumos",
}

# Prefijo (2 primeros dígitos del código de recurso) -> familia de gasto
FAMILIA_PREFIJOS = {
    11: 'GRH', 12: 'GRH',
    21: 'GCP',
    22: 'GGV',
    23: 'GGI',
    24: 'GGM',
    25: 'OGG',
    31: 'ST',
    41: 'MEI',
    51: 'ST',
    61: 'GPA', 62: 'GPA', 63: 'GPA', 64: 'GPA', 65: 'GPA', 66: 'GPA',
    71: 'STI', 72: 'STI', 73: 'STI', 74: 'STI', 75: 'STI',
}
FAMILIA_NOMBRE = {
    'GRH': 'GRH - Recursos Humanos',
    'GCP': 'GCP - Gastos Generales de Personal',
    'GGV': 'GGV - Vehículos y Equipos',
    'GGI': 'GGI - Bienes Inmuebles',
    'GGM': 'GGM - Bienes Muebles',
    'OGG': 'OGG - Otros Gastos Generales',
    'ST': 'ST - Servicios Tercerizados',
    'MEI': 'MEI - Materiales e Insumos',
    'GPA': 'GPA - Gasto Prestaciones Asociadas',
    'STI': 'STI - Servicios de Terceros (Estudios e Inversión)',
}

def familia_de_recurso(cod):
    prefijo = cod // 100
    key = FAMILIA_PREFIJOS.get(prefijo, 'OTR')
    return FAMILIA_NOMBRE.get(key, 'Otro (' + str(cod) + ')')

def nombre_recurso(cod):
    return RECURSO_NOMBRE.get(cod, f"Recurso {cod}")


RECURSOS_GGM = list(range(2401, 2412))
RECURSOS_OGG = list(range(2501, 2551))
RECURSOS_MEI = [4101, 4102, 4103, 4104, 4105, 4106]

# ============================================================================
# Familia ST - Servicios Tercerizados
# ============================================================================
# 29 tablas independientes. Cada tabla puede tener uno o más códigos de
# recurso (algunas comparten código con otras, ej. ST_22..ST_30 -> 5110;
# ST_16/ST_17 -> 5106 y 5109 a la vez). Por eso NO se asume un código fijo
# por archivo: se lee la columna "CÓDIGO RECURSO" de cada fila (igual que
# GGI_5 / GCP_5), detectando las columnas por NOMBRE de encabezado ya que
# no se conoce de antemano la estructura exacta de las 29 tablas.
ST_TABLE_TO_CODES = {
    "ST_3": [3101], "ST_4": [3102], "ST_5": [3103], "ST_6": [3104],
    "ST_7": [3105], "ST_8": [3106], "ST_9": [3107], "ST_10": [3108],
    "ST_11": list(range(3114, 3133)),  # 3114 a 3132
    "ST_12": [5101], "ST_13": [5102], "ST_14": [5105], "ST_15": [5106],
    "ST_16": [5106, 5109], "ST_17": [5106, 5109], "ST_18": [5107],
    "ST_19": [5114], "ST_20": [5108], "ST_21": [5109], "ST_22": [5110],
    "ST_23": [5110], "ST_25": [5110], "ST_27": [5110], "ST_28": [5110],
    "ST_29": [5110], "ST_30": [5110], "ST_31": [5111], "ST_32": [5112],
    "ST_33": [5114], "ST_34": [3133],
}
ST_TABLES = list(ST_TABLE_TO_CODES.keys())
RECURSOS_ST = sorted({c for codes in ST_TABLE_TO_CODES.values() for c in codes})

# ============================================================================
# Familia GPA - Gasto Prestaciones Asociadas
# ============================================================================
# 6 tablas (GPA_1 a GPA_6). Cada una trae su propio "CÓDIGO RECURSO" fijo
# (6101, 6201, 6301, 6401, 6501, 6601) y, a diferencia de GGM/OGG/MEI/ST, el
# 100% de su gasto se asigna SIEMPRE a un servicio regulado específico por
# tabla (no 1101, y sin necesidad de parametrizar servicios no regulados).
GPA_TABLE_TO_SERVICIO = {
    "GPA_1": 1201, "GPA_2": 1202, "GPA_3": 1203,
    "GPA_4": 1204, "GPA_5": 1205, "GPA_6": 1201,
}
GPA_TABLES = list(GPA_TABLE_TO_SERVICIO.keys())


def _normalizar_encabezado(texto):
    if texto is None:
        return ""
    txt = str(texto).upper().strip()
    for a, b in {"Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "Ñ": "N"}.items():
        txt = txt.replace(a, b)
    return " ".join(txt.split())


def _match_aproximado(palabra, texto_colapsado, umbral=0.82):
    """True si `palabra` aparece exacta en texto_colapsado, o si existe un
    tramo de largo similar muy parecido (tolera 1-2 errores de tipeo).
    Palabras muy cortas (<4 letras) solo se buscan de forma exacta, para
    evitar falsos positivos."""
    if palabra in texto_colapsado:
        return True
    if len(palabra) < 4:
        return False
    n = len(palabra)
    mejor = 0.0
    for i in range(0, max(1, len(texto_colapsado) - n + 1) + 3):
        trozo = texto_colapsado[i:i + n]
        if not trozo:
            continue
        score = difflib.SequenceMatcher(None, palabra, trozo).ratio()
        if score > mejor:
            mejor = score
    return mejor >= umbral


def _encontrar_columna(header_row, grupos_palabras_clave, evitar=None):
    """Busca una columna cuyo encabezado contenga TODAS las palabras clave
    de algún grupo (en cualquier orden, con o sin espacios entre ellas,
    tolerando errores de tipeo menores como "Recuso" en vez de "Recurso",
    o "CódigoRecurso" pegado sin espacio). Prueba los grupos en orden de
    prioridad y devuelve el índice del primer match. `evitar`: si el
    encabezado contiene alguna de estas palabras, se descarta la columna
    (para no confundir columnas parecidas, ej. no tomar "% Activado" como
    la columna de gasto)."""
    evitar = evitar or []
    normalizados = [_normalizar_encabezado(h) for h in header_row]
    colapsados = [n.replace(" ", "") for n in normalizados]

    for palabras in grupos_palabras_clave:
        for i, (norm, cole) in enumerate(zip(normalizados, colapsados)):
            if any(ev in norm for ev in evitar):
                continue
            if all(_match_aproximado(p, cole) for p in palabras):
                return i
    return None


def _a_numero(valor):
    """Convierte celdas numéricas 'sucias' (texto, '-', separadores de miles
    en formato chileno o estadounidense, vacías) a float de forma segura.
    Las tablas ST son cargadas por el usuario y pueden traer celdas con
    formato de texto en vez de número (ej. '-' como marcador de cero,
    '1.234.567' con puntos de miles, o '1.234,56' con coma decimal)."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, str):
        v = valor.strip()
        if v in ("", "-", "—", "–", "N/A", "n/a", "s/i", "S/I"):
            return 0.0
        v = v.replace("$", "").replace(" ", "")
        n_puntos = v.count(".")
        n_comas = v.count(",")
        if n_puntos >= 1 and n_comas >= 1:
            if v.rfind(",") > v.rfind("."):
                v = v.replace(".", "").replace(",", ".")  # chileno: 1.234.567,89
            else:
                v = v.replace(",", "")  # estadounidense: 1,234,567.89
        elif n_comas >= 1:
            if n_comas >= 2:
                v = v.replace(",", "")
            else:
                entero, dec = v.split(",")
                if entero.lstrip("-").isdigit() and len(dec) <= 2:
                    v = entero + "." + dec  # coma decimal (ej. "1234,5")
                else:
                    v = entero + dec
        elif n_puntos >= 2:
            v = v.replace(".", "")  # miles chilenos "1.234.567"
        elif n_puntos == 1:
            entero, dec = v.split(".")
            if entero.lstrip("-").isdigit() and len(dec) == 3 and len(entero.lstrip("-")) <= 3:
                v = entero + dec  # miles "1.234" (no decimal, montos en pesos)
        try:
            return float(v)
        except ValueError:
            return 0.0
    return 0.0


def leer_tabla_st(file_bytes):
    """Lee una tabla ST_x detectando columnas por nombre de encabezado.
    Devuelve lista de (cod_recurso, monto_activado, total_gasto_no_activado).
    Lanza ValueError si no logra identificar las columnas necesarias."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    header_row = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    idx_recurso = _encontrar_columna(header_row, [["CODIGO", "RECURSO"], ["COD", "RECURSO"]])
    idx_activado = _encontrar_columna(header_row, [["MONTO", "ANUAL", "ACTIVADO"], ["MONTO", "ACTIVADO"]])
    idx_gasto = _encontrar_columna(
        header_row,
        [["TOTAL", "GASTO", "ANUAL"], ["GASTO", "ANUAL", "NO", "ACTIVADO"], ["TOTAL", "GASTO"], ["GASTO"]],
        evitar=["ACTIVADO", "%"],
    )

    faltantes = []
    if idx_recurso is None:
        faltantes.append("CÓDIGO RECURSO")
    if idx_activado is None:
        faltantes.append("MONTO ANUAL ACTIVADO")
    if idx_gasto is None:
        faltantes.append("TOTAL GASTO ANUAL")
    if faltantes:
        raise ValueError(f"No se identificaron las columnas: {', '.join(faltantes)}")

    filas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        cod_recurso = row[idx_recurso]
        if cod_recurso is None:
            continue
        cod_recurso = _a_numero(cod_recurso)
        cod_recurso = int(cod_recurso) if cod_recurso == int(cod_recurso) else cod_recurso
        filas.append((cod_recurso, _a_numero(row[idx_activado]), _a_numero(row[idx_gasto])))
    return filas


def identificar_tabla(nombre_archivo, tablas_conocidas):
    """Extrae el nombre de tabla (ej. 'ST_12', 'GPA_3') de un archivo subido,
    sin importar mayúsculas, extensión o sufijos. Devuelve None si no
    matchea ninguna tabla conocida de la lista dada."""
    base = nombre_archivo.upper().replace(".XLSX", "").replace(".XLS", "")
    for tabla in tablas_conocidas:
        if base == tabla or base.startswith(tabla + "_") or base.startswith(tabla + "-"):
            return tabla
    return None


def identificar_tabla_st(nombre_archivo):
    return identificar_tabla(nombre_archivo, ST_TABLES)


def identificar_tabla_gpa(nombre_archivo):
    return identificar_tabla(nombre_archivo, GPA_TABLES)


def read_rows(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    return [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None]


def familia(cod_servicio):
    return cod_servicio // 100


def redondear_para_sumar_100(valores, decimales=4):
    """Redondea una lista de fracciones que en teoría suman ~1.0 usando el
    método de mayor residuo (largest remainder / Hare-Niemeyer), de forma
    que la suma de los valores redondeados sea EXACTAMENTE 1.0 al número de
    decimales indicado. Evita el arrastre de error que ocurre al redondear
    cada valor de forma independiente (ej. una suma final de 99.99% o
    100.01% en vez de 100.00%)."""
    if not valores:
        return []
    escala = 10 ** decimales
    total_objetivo = round(sum(valores) * escala)
    escalados = [v * escala for v in valores]
    pisos = [int(x) for x in escalados]
    residuos = [x - p for x, p in zip(escalados, pisos)]
    suma_pisos = sum(pisos)
    faltante = total_objetivo - suma_pisos
    orden = sorted(range(len(valores)), key=lambda i: residuos[i], reverse=True)
    ajustados = list(pisos)
    for i in range(faltante):
        ajustados[orden[i % len(valores)]] += 1
    return [a / escala for a in ajustados]


def fusionar_cola_larga(items, piso=0.01, id_fusion="-1"):
    """Si hay demasiadas sub-claves para que todas alcancen el piso mínimo
    individualmente (n * piso >= 1.0), fusiona las de MENOR monto bajo
    id_fusion ('-1', el código que el maestro SISS reserva para 'sin cliente
    específico'), dejando como entradas individuales solo las que sí pueden
    alcanzar el piso con margen para variación proporcional.

    items: [(sub, monto), ...] donde 'sub' puede ser un valor simple (ej. el
    ID cliente) o una tupla (ej. (cod_servicio, id_cliente)). En este último
    caso, la fusión agrupa la 'cola' POR EL RESTO DE LA TUPLA (ej. por
    servicio) y reemplaza solo el último componente por id_fusion, de forma
    que cada fila fusionada siga siendo válida (ej. conserva su servicio).

    Devuelve (nuevos_items, se_fusiono)."""
    n = len(items)
    max_individuales = int(1.0 / piso) - 1
    if n <= max_individuales:
        return items, False

    items_ordenados = sorted(items, key=lambda x: x[1], reverse=True)
    individuales = items_ordenados[:max_individuales - 1]
    cola = items_ordenados[max_individuales - 1:]

    nuevos = list(individuales)
    if cola and isinstance(cola[0][0], tuple):
        # Agrupar la cola por el resto de la tupla (ej. por servicio),
        # reemplazando solo el último componente (ej. el cliente) por id_fusion.
        cola_por_grupo = defaultdict(float)
        for sub, monto in cola:
            clave_grupo = sub[:-1]
            cola_por_grupo[clave_grupo] += monto
        existentes = {s[0][:-1]: i for i, s in enumerate(nuevos) if s[0][-1] == id_fusion}
        for clave_grupo, monto in cola_por_grupo.items():
            nuevo_sub = clave_grupo + (id_fusion,)
            if clave_grupo in existentes:
                idx = existentes[clave_grupo]
                nuevos[idx] = (nuevos[idx][0], nuevos[idx][1] + monto)
            else:
                nuevos.append((nuevo_sub, monto))
    else:
        monto_cola = sum(m for _, m in cola)
        fusion_existente = False
        for i, (cid, m) in enumerate(nuevos):
            if cid == id_fusion:
                nuevos[i] = (cid, m + monto_cola)
                fusion_existente = True
                break
        if not fusion_existente:
            nuevos.append((id_fusion, monto_cola))
    return nuevos, True


def aplicar_piso_minimo(fracciones, piso=0.01):
    """Ajusta una lista de fracciones (que suman ~1.0) para que NINGUNA
    quede por debajo de 'piso' (ej. 0.01 = 1%), redistribuyendo
    proporcionalmente el resto entre las demás ("water-filling"). Se usa en
    CYG_9 para cumplir la regla del maestro SISS de no declarar registros
    con % de dedicación igual a 0 (ni prácticamente 0), sin dejar de listar
    a todos los clientes que efectivamente reciben una porción del gasto.
    Si hay demasiadas entradas para que todas alcancen el piso (n*piso >= 1),
    se reparte equitativamente entre todas como mejor esfuerzo."""
    n = len(fracciones)
    if n == 0:
        return []
    if n * piso >= 1.0:
        return [1.0 / n] * n

    resultado = [0.0] * n
    restantes = list(range(n))
    masa_restante = 1.0

    while True:
        suma_raw = sum(fracciones[i] for i in restantes)
        candidatos_bajo_piso = []
        for i in restantes:
            val = (fracciones[i] / suma_raw * masa_restante) if suma_raw > 0 else (masa_restante / len(restantes))
            if val < piso:
                candidatos_bajo_piso.append(i)

        if not candidatos_bajo_piso:
            for i in restantes:
                resultado[i] = (fracciones[i] / suma_raw * masa_restante) if suma_raw > 0 else (masa_restante / len(restantes))
            break

        for i in candidatos_bajo_piso:
            resultado[i] = piso
            restantes.remove(i)
            masa_restante -= piso

        if not restantes:
            break
        if len(restantes) * piso >= masa_restante:
            for i in restantes:
                resultado[i] = masa_restante / len(restantes)
            break

    return resultado


def normalizar_con_piso_minimo(items_por_clave, piso=0.01, decimales=4):
    """Como _normalizar_por_clave, pero garantiza que ningún % quede por
    debajo de 'piso' (ej. 0.01 = 1%). Si hay demasiadas sub-claves (ej.
    clientes) para que todas alcancen el piso, primero fusiona la cola de
    menor monto bajo el identificador '-1' (fusionar_cola_larga), y luego
    aplica el piso mínimo (aplicar_piso_minimo) sobre el conjunto resultante.
    Devuelve también, por clave, si hubo fusión (para poder avisar)."""
    resultado = {}
    hubo_fusion = {}
    for clave, items in items_por_clave.items():
        total = sum(g for _, g in items)
        if total == 0:
            continue
        items_fusionados, se_fusiono = fusionar_cola_larga(items, piso=piso, id_fusion="-1")
        hubo_fusion[clave] = se_fusiono
        fracciones_raw = [g / total for _, g in items_fusionados]
        fracciones_piso = aplicar_piso_minimo(fracciones_raw, piso)
        pcts = redondear_para_sumar_100(fracciones_piso, decimales)
        resultado[clave] = [(sub, pct, pct * total) for (sub, _), pct in zip(items_fusionados, pcts)]
    return resultado, hubo_fusion


def procesar_familia_plana(agg, EMPRESA, PERIODO, ANIO, SECTOR, tablas_specs, params):
    by_recurso = defaultdict(lambda: [0.0, 0.0])
    for filas, idx_r, idx_a, idx_g in tablas_specs:
        for r in filas:
            cod_recurso = r[idx_r]
            monto_act = _a_numero(r[idx_a])
            total_gasto = _a_numero(r[idx_g])
            by_recurso[cod_recurso][0] += total_gasto
            by_recurso[cod_recurso][1] += monto_act

    for cod_recurso, (gasto_no_act, monto_act) in by_recurso.items():
        overrides = params.get(cod_recurso, [])
        pct_reg = 1.0 - sum(p for _, p in overrides)
        fam = familia(1101)
        k = (EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam)
        agg[k][0] += gasto_no_act * pct_reg
        agg[k][1] += monto_act  # monto activado SIEMPRE 100% al servicio regulado por defecto
        for cod_serv_noreg, pct in overrides:
            fam2 = familia(cod_serv_noreg)
            k2 = (EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam2)
            agg[k2][0] += gasto_no_act * pct
            # el monto activado no se traspasa a servicios no regulados
    return by_recurso


def safe_read_rows(file_bytes, etiqueta, avisos):
    """Como read_rows, pero tolera archivo ausente (None) o con error de
    lectura: registra un aviso y devuelve lista vacía en vez de fallar."""
    if file_bytes is None:
        avisos.append(f"⚠️ Falta la tabla **{etiqueta}** — se excluye del cálculo (esa familia/recurso queda en 0 o incompleto).")
        return []
    try:
        filas = read_rows(file_bytes)
    except Exception as e:
        avisos.append(f"⚠️ No se pudo leer **{etiqueta}** ({e}) — se excluye del cálculo.")
        return []
    if len(filas) == 0:
        avisos.append(f"ℹ️ La tabla **{etiqueta}** está vacía (sin filas de datos).")
    return filas


def build_rep2(fb, ggm_params, ogg_params, mei_params, st_files, st_params, gpa_files):
    avisos = []

    grh8 = safe_read_rows(fb.get("grh8"), "GRH_8", avisos)
    grh11 = safe_read_rows(fb.get("grh11"), "GRH_11", avisos)
    gcp4 = safe_read_rows(fb.get("gcp4"), "GCP_4", avisos)
    gcp5 = safe_read_rows(fb.get("gcp5"), "GCP_5", avisos)
    ggv4 = safe_read_rows(fb.get("ggv4"), "GGV_4", avisos)
    ggv5 = safe_read_rows(fb.get("ggv5"), "GGV_5", avisos)
    ggi5 = safe_read_rows(fb.get("ggi5"), "GGI_5", avisos)
    ggm_tablas = [safe_read_rows(fb.get(f"ggm{i}"), f"GGM_{i}", avisos) for i in range(1, 6)]
    ogg5 = safe_read_rows(fb.get("ogg5"), "OGG_5", avisos)
    mei1 = safe_read_rows(fb.get("mei1"), "MEI_1", avisos)
    mei2 = safe_read_rows(fb.get("mei2"), "MEI_2", avisos)
    mei3 = safe_read_rows(fb.get("mei3"), "MEI_3", avisos)
    mei4 = safe_read_rows(fb.get("mei4"), "MEI_4", avisos)

    agg = defaultdict(lambda: [0.0, 0.0])

    # EMPRESA/PERIODO/AÑO/SECTOR: se toman del primer archivo disponible,
    # ya que si falta GRH_8 no podemos asumir esos valores desde ahí.
    EMPRESA = PERIODO = ANIO = SECTOR = None
    for candidato in [grh8, gcp4, ggv4, ggi5, ogg5] + ggm_tablas + [mei1, mei2, mei3, mei4]:
        if candidato:
            EMPRESA, PERIODO, ANIO, SECTOR = candidato[0][0], candidato[0][1], candidato[0][2], candidato[0][3]
            break
    if EMPRESA is None:
        for filas_st in st_files.values():
            if filas_st:
                EMPRESA, PERIODO, ANIO, SECTOR = filas_st[0][0], filas_st[0][1], filas_st[0][2], filas_st[0][3]
                break

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

    # --- ST: Servicios Tercerizados (29 tablas, lectura por nombre de columna) ---
    st_by_recurso = defaultdict(lambda: [0.0, 0.0])
    tablas_st_faltantes = [t for t in ST_TABLES if t not in st_files]
    if tablas_st_faltantes:
        avisos.append(
            f"⚠️ Faltan {len(tablas_st_faltantes)} tabla(s) ST: {', '.join(tablas_st_faltantes)} "
            "— se excluyen del cálculo (sus códigos de recurso quedan en 0)."
        )
    for nombre_tabla, filas_st in st_files.items():
        if len(filas_st) == 0:
            avisos.append(f"ℹ️ La tabla **{nombre_tabla}** está vacía (sin filas de datos).")
            continue
        for cod_recurso, monto_act, total_gasto in filas_st:
            overrides = st_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)
            fam = familia(1101)
            k = (EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam)
            agg[k][0] += total_gasto * pct_reg
            agg[k][1] += monto_act  # monto activado SIEMPRE 100% al servicio regulado por defecto
            st_by_recurso[cod_recurso][0] += total_gasto * pct_reg
            st_by_recurso[cod_recurso][1] += monto_act
            for cod_serv_noreg, pct in overrides:
                fam2 = familia(cod_serv_noreg)
                k2 = (EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam2)
                agg[k2][0] += total_gasto * pct
                # el monto activado no se traspasa a servicios no regulados
                st_by_recurso[cod_recurso][0] += total_gasto * pct

    # --- GPA: Gasto Prestaciones Asociadas (6 tablas, 100% a servicio fijo por tabla) ---
    gpa_by_recurso = defaultdict(lambda: [0.0, 0.0])
    gpa_detalle = []  # (tabla, cod_recurso, servicio_asignado, gasto, activado) para reporte
    tablas_gpa_faltantes = [t for t in GPA_TABLES if t not in gpa_files]
    if tablas_gpa_faltantes:
        avisos.append(
            f"⚠️ Faltan {len(tablas_gpa_faltantes)} tabla(s) GPA: {', '.join(tablas_gpa_faltantes)} "
            "— se excluyen del cálculo (sus códigos de recurso quedan en 0)."
        )
    for nombre_tabla, filas_gpa in gpa_files.items():
        if len(filas_gpa) == 0:
            avisos.append(f"ℹ️ La tabla **{nombre_tabla}** está vacía (sin filas de datos).")
            continue
        servicio_fijo = GPA_TABLE_TO_SERVICIO[nombre_tabla]
        fam = familia(servicio_fijo)
        tabla_g, tabla_a = 0.0, 0.0
        recursos_en_tabla = set()
        for cod_recurso, monto_act, total_gasto in filas_gpa:
            k = (EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam)
            agg[k][0] += total_gasto
            agg[k][1] += monto_act
            gpa_by_recurso[cod_recurso][0] += total_gasto
            gpa_by_recurso[cod_recurso][1] += monto_act
            tabla_g += total_gasto
            tabla_a += monto_act
            recursos_en_tabla.add(cod_recurso)
        recurso_label = ", ".join(str(c) for c in sorted(recursos_en_tabla))
        gpa_detalle.append((nombre_tabla, recurso_label, servicio_fijo, tabla_g, tabla_a))

    # --- Normalizar % por recurso (ajustado para sumar EXACTAMENTE 100%) ---
    grupos_por_recurso = defaultdict(list)  # cod_recurso -> [(key_completa, g, a), ...]
    for key, (g, a) in agg.items():
        cod_recurso = key[4]
        grupos_por_recurso[cod_recurso].append((key, g, a))

    final_rows = []
    for cod_recurso, items in grupos_por_recurso.items():
        items = [(k, g, a) for k, g, a in items if not (abs(g) < 1e-9 and abs(a) < 1e-9)]
        if not items:
            continue
        total_g = sum(g for _, g, _ in items)
        total_a = sum(a for _, _, a in items)

        idx_g_pos = [i for i, (_, g, _) in enumerate(items) if abs(g) > 1e-9]
        fracciones_g = [items[i][1] / total_g for i in idx_g_pos] if total_g else []
        pcts_g_corr = redondear_para_sumar_100(fracciones_g, 4) if fracciones_g else []
        pct_g_final = {i: 0.0 for i in range(len(items))}
        for idx, pct in zip(idx_g_pos, pcts_g_corr):
            pct_g_final[idx] = pct

        idx_a_pos = [i for i, (_, _, a) in enumerate(items) if abs(a) > 1e-9]
        fracciones_a = [items[i][2] / total_a for i in idx_a_pos] if total_a else []
        pcts_a_corr = redondear_para_sumar_100(fracciones_a, 4) if fracciones_a else []
        pct_a_final = {i: 0.0 for i in range(len(items))}
        for idx, pct in zip(idx_a_pos, pcts_a_corr):
            pct_a_final[idx] = pct

        for i, (key, g, a) in enumerate(items):
            empresa, periodo, anio, sector, _, fam = key
            final_rows.append([
                empresa, periodo, anio, sector, cod_recurso, fam,
                round(pct_g_final[i] * 100, 2), round(g, 2),
                round(pct_a_final[i] * 100, 2), round(a, 2),
            ])

    final_rows.sort(key=lambda r: (r[4], r[5]))

    # --- Validaciones (solo entre familias con datos disponibles) ---
    recursos_grh = set(r[6] for r in grh8)
    recursos_gcp = set(r[6] for r in gcp4)
    recursos_ggv = set(r[5] for r in ggv4)
    recursos_ggi = set(r[5] for r in ggi5)
    recursos_ggm = set(ggm_by_recurso.keys())
    recursos_ogg = set(ogg_by_recurso.keys())
    recursos_mei = set(mei_by_recurso.keys())
    recursos_st = set(st_by_recurso.keys())
    recursos_gpa = set(gpa_by_recurso.keys())

    def validar(recursos, sum_gasto_fuente, sum_act_fuente):
        sum_gasto_rep2 = sum(r[7] for r in final_rows if r[4] in recursos)
        sum_act_rep2 = sum(r[9] for r in final_rows if r[4] in recursos)
        return {
            "diff_gasto": sum_gasto_rep2 - sum_gasto_fuente,
            "diff_act": sum_act_rep2 - sum_act_fuente,
        }

    checks = {}
    if grh8:
        checks["GRH"] = validar(recursos_grh, sum(r[10] for r in grh8), sum(r[8] for r in grh8))
    if gcp4:
        checks["GCP"] = validar(recursos_gcp, sum(r[10] for r in gcp4), sum(r[8] for r in gcp4))
    if ggv4:
        checks["GGV"] = validar(recursos_ggv, sum(r[9] for r in ggv4), sum(r[7] for r in ggv4))
    if ggi5:
        checks["GGI"] = validar(recursos_ggi, sum(r[6] for r in ggi5), 0)
    if ggm_by_recurso:
        checks["GGM"] = validar(recursos_ggm, sum(v[0] for v in ggm_by_recurso.values()), sum(v[1] for v in ggm_by_recurso.values()))
    if ogg_by_recurso:
        checks["OGG"] = validar(recursos_ogg, sum(v[0] for v in ogg_by_recurso.values()), sum(v[1] for v in ogg_by_recurso.values()))
    if mei_by_recurso:
        checks["MEI"] = validar(recursos_mei, sum(v[0] for v in mei_by_recurso.values()), sum(v[1] for v in mei_by_recurso.values()))
    if st_by_recurso:
        checks["ST"] = validar(recursos_st, sum(v[0] for v in st_by_recurso.values()), sum(v[1] for v in st_by_recurso.values()))
    if gpa_by_recurso:
        checks["GPA"] = validar(recursos_gpa, sum(v[0] for v in gpa_by_recurso.values()), sum(v[1] for v in gpa_by_recurso.values()))

    familia_map = {}
    for c in recursos_grh: familia_map[c] = "GRH - Gastos Recursos Humanos"
    for c in recursos_gcp: familia_map[c] = "GCP - Gastos Generales de Personal"
    for c in recursos_ggv: familia_map[c] = "GGV - Gastos Generales Vehículos y Equipos"
    for c in recursos_ggi: familia_map[c] = "GGI - Gastos Generales Bienes Inmuebles"
    for c in recursos_ggm: familia_map[c] = "GGM - Gastos Generales Bienes Muebles"
    for c in recursos_ogg: familia_map[c] = "OGG - Otros Gastos Generales"
    for c in recursos_mei: familia_map[c] = "MEI - Materiales e Insumos"
    for c in recursos_st: familia_map[c] = "ST - Servicios Tercerizados"
    for c in recursos_gpa: familia_map[c] = "GPA - Gasto Prestaciones Asociadas"

    by_recurso_planas = {"GGM": ggm_by_recurso, "OGG": ogg_by_recurso, "MEI": mei_by_recurso, "ST": st_by_recurso}
    return final_rows, checks, familia_map, by_recurso_planas, gpa_detalle, avisos


def build_excel(final_rows, familia_map, by_recurso_planas, params_by_familia, gpa_detalle=None, avisos=None, template_bytes=None):

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
                cell.number_format = "0.00"
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
        "MEI - Materiales e Insumos", "ST - Servicios Tercerizados",
        "GPA - Gasto Prestaciones Asociadas",
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
                cell.number_format = "0.00"
    ws3.column_dimensions["A"].width = 10
    ws3.column_dimensions["B"].width = 16
    ws3.column_dimensions["C"].width = 22
    ws3.column_dimensions["D"].width = 42
    ws3.column_dimensions["E"].width = 55
    ws3.column_dimensions["F"].width = 12

    # Asignación GPA (fija, sin parametrización de servicios no regulados)
    if gpa_detalle:
        ws3b = wb.create_sheet("Asignacion_GPA")
        ws3b.append(["Tabla", "CÓDIGO(S) RECURSO", "SERVICIO REGULADO ASIGNADO (100%)", "GASTO ANUAL", "MONTO ACTIVADO"])
        for c in range(1, 6):
            cell = ws3b.cell(row=1, column=c)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws3b.row_dimensions[1].height = 30
        for tabla, cod_recurso, servicio, g, a in sorted(gpa_detalle):
            ws3b.append([tabla, cod_recurso, servicio, round(g, 2), round(a, 2)])
        for r in range(2, ws3b.max_row + 1):
            for c in range(1, 6):
                cell = ws3b.cell(row=r, column=c)
                cell.font = data_font
                cell.border = border
                if c in (4, 5):
                    cell.number_format = '#,##0;(#,##0);"-"'
        ws3b.column_dimensions["A"].width = 12
        ws3b.column_dimensions["B"].width = 16
        ws3b.column_dimensions["C"].width = 32
        ws3b.column_dimensions["D"].width = 18
        ws3b.column_dimensions["E"].width = 18

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

    # Avisos de carga (tablas faltantes o vacías, generación parcial)
    if avisos:
        ws5 = wb.create_sheet("Avisos_Carga")
        ws5.append(["Aviso"])
        cell = ws5.cell(row=1, column=1)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        for aviso in avisos:
            ws5.append([aviso.replace("**", "")])
        for r in range(2, ws5.max_row + 1):
            cell = ws5.cell(row=r, column=1)
            cell.font = data_font
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
        ws5.column_dimensions["A"].width = 110

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
        sel_pct = st.number_input(
            "% asignado", min_value=0.0, max_value=100.0, value=10.0,
            step=0.0000000001, format="%.10f", key=f"sel_pct_{session_key}",
            help="Puedes ingresar hasta 10 decimales para minimizar la diferencia de redondeo en el monto resultante (con montos de cientos de millones, 6 decimales aún puede dejar una diferencia de $10-30; con 9-10 decimales queda en centavos). La columna '% ASIGNADO' del reporte final siempre se mostrará con 2 decimales (formato regulatorio), pero el MONTO se calcula con la precisión completa que ingreses aquí.",
        )
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
            c3.write(f"{pct:.9%}")
            if c4.button("Quitar", key=f"del_{session_key}_{i}"):
                st.session_state[session_key].pop(i)
                st.rerun()
    else:
        st.caption(f"Sin parametrización: 100% de todos los recursos {nombre} va al servicio 1101.")

    params = defaultdict(list)
    for cod_r, cod_s, pct in st.session_state[session_key]:
        params[cod_r].append((cod_s, pct))
    return dict(params)



# ============================================================================
# REP_3 - Costos y Gastos No Activados de Servicios Regulados - Proceso
# ============================================================================
# Códigos de actividad válidos según MAE_1 del Maestro SISS (617 códigos).
# Se usa para validar que los códigos de actividad de las tablas ST (y
# alertar si alguno no corresponde a una actividad SISS reconocida).
CODIGOS_ACTIVIDAD_VALIDOS = frozenset([
    1010101, 1010201, 1010301, 1010302, 1010303, 1010304, 1010305, 1010306, 1010307, 1020101, 1020201, 1020301, 1020302, 1020303, 1020304,
    1030101, 1030201, 1030202, 1030301, 1030302, 1030303, 1030304, 1030401, 1030402, 1030403, 1030404, 1030501, 1030502, 1030503, 1030504,
    1030505, 1030601, 1030602, 1030603, 1030604, 1030605, 1030701, 1030702, 1030703, 1030704, 1030705, 1030801, 1030802, 1030803, 1030901,
    1030902, 1030903, 1031001, 1031002, 1031101, 1031201, 1031202, 1040101, 1040201, 1040301, 1040302, 1040303, 1040304, 1040305, 1040306,
    1040307, 1040308, 1040309, 1040401, 1050101, 1050102, 1060101, 1060102, 1060201, 1060202, 1060203, 1060204, 1060205, 1060301, 1060302,
    2010101, 2010201, 2010301, 2010302, 2010303, 2010304, 2010305, 2010306, 2010307, 2020101, 2020102, 2020103, 2020201, 2020301, 2030101,
    2030201, 2030301, 2040101, 2040102, 2040103, 2040104, 2040201, 2040301, 2040401, 2040402, 2040403, 2040404, 2040405, 2040501, 2040601,
    2040701, 2040702, 2040703, 2040704, 2040705, 2040801, 2040802, 2040901, 2040902, 2040903, 2040904, 2041001, 3010101, 3010201, 3010301,
    3010401, 3010501, 3010601, 3010701, 3010801, 3010901, 3011001, 3011101, 3011201, 3011301, 3020101, 3020201, 3020301, 3020401, 3020501,
    3030101, 3030201, 3030301, 3030302, 3030303, 3030304, 3030401, 3030402, 3030403, 3030404, 3030501, 3030502, 3030503, 3030504, 3030505,
    3030601, 3030602, 3030603, 3030701, 3030702, 3030703, 3030801, 3030802, 3030901, 3031001, 3031002, 3040101, 3040201, 3040301, 3040401,
    3040501, 3050101, 3050201, 3050301, 3050302, 3050303, 3050401, 3050402, 3050501, 3050502, 3060101, 3060201, 3060301, 3060302, 3060303,
    3060304, 3060401, 3060402, 3060403, 3060404, 3060501, 3060502, 3060503, 3060504, 3070101, 3070201, 3070301, 3070302, 3080101, 3080102,
    3080103, 3090101, 3090201, 3090301, 3090302, 3090303, 3090401, 3090402, 3090403, 3090501, 3090502, 3090503, 3090504, 3100101, 3100201,
    3100301, 3100302, 3100303, 3100304, 3100305, 3100401, 3100402, 3100403, 3100404, 3100405, 3100501, 3100502, 3100503, 3100504, 3110101,
    3110201, 3110301, 3110302, 3110303, 3110304, 3120101, 3120201, 3120301, 3120401, 3120402, 3120403, 3120404, 3120405, 3120501, 3120601,
    3130101, 3130201, 3130301, 3130401, 3130501, 4010101, 4010102, 4010103, 4010104, 4010105, 4010201, 4010202, 4010203, 4010204, 4010205,
    4010206, 4010207, 4010301, 4010302, 4010303, 4010304, 4010401, 4010402, 4010403, 4010404, 4010405, 4010406, 4010407, 4010408, 4020101,
    4020102, 4020103, 4020201, 4020202, 4020203, 4020301, 4020302, 4020303, 4020304, 4020305, 4020306, 4020307, 4020308, 4020309, 4020310,
    4020311, 4020312, 4020313, 4020314, 4020315, 4020316, 4020317, 4020318, 4020319, 4020320, 4020321, 4020322, 4020323, 4020324, 4020325,
    4020326, 4020327, 4020401, 4020402, 4020403, 4020501, 4020502, 4020503, 4020601, 4020602, 4030101, 4030102, 4030103, 4030104, 4030105,
    4030106, 4030107, 4030108, 4030201, 4030202, 4030203, 4030204, 4030205, 4030206, 4030207, 4030208, 4030209, 4030210, 4030211, 4030301,
    4030302, 4030303, 4030304, 4030305, 4030306, 4030307, 4030308, 4030309, 4030401, 4030402, 4030403, 4030404, 5010101, 5010102, 5010103,
    5010104, 5010105, 5010106, 5010107, 5010108, 5010201, 5010202, 5010203, 5010204, 5010205, 5010206, 5010301, 5010302, 5010303, 5010304,
    5010305, 5010306, 5010307, 5010308, 5010401, 5010402, 5010403, 5010404, 5010405, 5010406, 5010407, 5010501, 5010502, 5010503, 5010504,
    5010505, 5010506, 5010601, 5010602, 5010603, 5010604, 5010605, 5010606, 5010607, 5010608, 5020101, 5020102, 5020103, 5020104, 5020105,
    5020106, 5020107, 5020108, 5020109, 5020110, 5020111, 5020112, 5020113, 5020114, 5020115, 5020201, 5020202, 5020203, 5020204, 5020205,
    5020206, 5020207, 5020208, 5020209, 5020210, 5020211, 5020212, 5020301, 5020302, 5020303, 5020304, 5020305, 5020306, 5030101, 5030102,
    5030103, 5030104, 5030105, 5030106, 5030107, 5030108, 5030109, 5030110, 5030111, 5030112, 5030113, 5030114, 5030115, 5030116, 5030117,
    5030118, 5030119, 5030120, 5030121, 5030122, 5030123, 5030124, 5030125, 5030201, 5030202, 5030203, 5030204, 5030205, 5030206, 5030207,
    5030208, 5030209, 5030210, 5030211, 5030212, 5030213, 5040101, 5040102, 5040103, 5040104, 5040105, 5040106, 5040107, 5040201, 5040202,
    5040203, 5040204, 5040205, 5040206, 5040207, 5040208, 5040209, 5040210, 5040301, 5040302, 5040303, 5040304, 5050101, 5050102, 5050103,
    5050104, 5050105, 5050106, 5050107, 5050108, 5050109, 5050110, 5050111, 5050112, 5050113, 5050114, 5050115, 5050116, 5060101, 5060102,
    5060103, 5060104, 5060105, 5060106, 5060107, 5060201, 5060202, 5060203, 5060204, 5060301, 5060302, 5060303, 5060304, 5060401, 5060402,
    5060403, 5060404, 5060405, 5060406, 5060501, 5060502, 5060503, 5060504, 5060505, 5060506, 5060507, 5060508, 5060509, 6010101, 6010102,
    6010103, 6010104, 6010105, 6010106, 6010107, 6010108, 6010109, 6010110, 6010111, 6010112, 6010113, 6010114, 6010201, 6010202, 6010203,
    6010204, 6010301, 6010302, 6010303, 6010304, 6010305, 6010306, 6010401, 6010402, 6010403, 6010404, 6010501, 6010502, 6010503, 6010504,
    6010505, 6010506, 6010507, 6010508, 6010509, 6010601, 6010602, 7010101, 7010102, 7010103, 7010104, 7010105, 7010201, 7010202, 7010203,
    7010204, 7010205, 7010206, 7010207, 7010301, 7010302, 7010303, 7010304, 7010305, 7020101, 7020102, 7020103, 7020104, 7020105, 7020106,
    7030101, 7030102, 7030103, 7030104, 7030105, 7030106, 7030201, 7030202, 7030203, 7030204, 7030301, 7030302, 7030303, 7030304, 7040101,
    7040102, 7040103, 7040104, 7040105, 7040201, 7040202, 7040203, 7040204, 7040205, 7050101, 7050102, 7050103, 7050201, 7050202, 7050301,
    7050302, 7060101,
])

# Catálogo de procesos (para nombrar los códigos de proceso en reportes)
PROCESO_NOMBRE = {
    101: "Operación Infraestructura de Apoyo AP", 102: "Gestión de Recursos Hídricos",
    103: "Operación Infraestructura de Capacidad AP", 104: "Operación Redes y Conducciones AP",
    105: "Operación Infraestructura de Capacidad AP", 106: "Operación Redes y Conducciones AP",
    201: "Operación Infraestructura de Apoyo AS", 202: "Operación Infraestructura de Capacidad AS",
    203: "Operación Redes y Conducciones AS", 204: "Operación Infraestructura de Capacidad AS",
    301: "Mantención Preventiva de Infraestructura de Apoyo", 302: "Mantención Correctiva de Infraestructura de Apoyo",
    303: "Mantención Preventiva de Infraestructura de Capacidad AP", 304: "Mantención Correctiva de Infraestructura de Capacidad AP",
    305: "Mantención Preventiva Redes y Conducciones AP", 306: "Mantención Correctiva Redes y Conducciones AP",
    307: "Mantención Preventiva Conexiones AP", 308: "Mantención Correctiva Conexiones AP",
    309: "Mantención Preventiva Redes y Conducciones AS", 310: "Mantención Correctiva Redes y Conducciones AS",
    311: "Mantención Correctiva Conexiones AS", 312: "Mantención Preventiva de Infraestructura de Capacidad AS",
    313: "Mantención Correctiva de Infraestructura de Capacidad AS",
    401: "Ciclo de Recaudación", 402: "Gestión Comercial", 403: "Gestión de Incorporación de Clientes",
    501: "Dirección Superior", 502: "Administración y Finanzas", 503: "Recursos Humanos",
    504: "Abastecimiento y Servicios Generales", 505: "Informática", 506: "Planificación y Desarrollo",
    601: "Prestaciones Asociadas",
    701: "Planificación y Diseño", 702: "Construcción de Obras", 703: "Inspección Técnica de Obras",
    704: "Telemetría, Comunicación y Sistemas de Información", 705: "Adquisición Bienes Muebles e Inmuebles",
    706: "Adquisición de Derechos de Agua",
}


def cod_proceso_de_actividad(cod_actividad):
    """Código de Proceso = primeros 3 dígitos del Código Actividad (7 dígitos),
    según MAE_1 del Maestro SISS."""
    return int(str(int(cod_actividad)).zfill(7)[:3])


def nombre_proceso(cod_proceso):
    return PROCESO_NOMBRE.get(cod_proceso, f"Proceso {cod_proceso}")


# --- Defaults de proceso para familias SIN tabla de apertura por actividad ---
# GGM: costos de informática/telecomunicaciones -> Informática (505);
#      materiales de oficina/laboratorio -> Abastecimiento y Servicios Generales (504)
DEFAULT_PROCESO_GGM = {
    2401: 505, 2402: 505, 2403: 505, 2404: 505, 2405: 505,
    2406: 505, 2407: 505, 2408: 505, 2409: 505,
    2410: 504, 2411: 504,
}
# OGG: gastos de Directorio -> Dirección Superior (501); resto -> Administración y Finanzas (502)
DEFAULT_PROCESO_OGG = {2501: 501, 2502: 501}
# Todo recurso OGG no listado explícitamente usa 502 por defecto (se resuelve en el código)
DEFAULT_PROCESO_OGG_FALLBACK = 502

# MEI_1 (4101 Productos químicos): sin columna de actividad. Se usan para
# tratamiento de agua potable (cloración/fluoración, proceso 103) y de aguas
# servidas (pretratamiento/tratamiento, proceso 204). Referencia empírica:
# la distribución real de MEI_2 (Energía Eléctrica, que SÍ trae actividad)
# entre estos mismos procesos de tratamiento es 103≈51% / 204≈23% (proporción
# ≈70/30 entre ambos) — se usa esa proporción como default razonable.
DEFAULT_PROCESO_MEI1 = {4101: [(103, 0.70), (204, 0.30)]}

# GPA: cada tabla corresponde 1 a 1 a una "Prestación Asociada" específica
# (Control de Riles, Mantención de Grifos, Corte y Reposición, Revisión de
# Proyectos, Verificación de Medidores, Otras Prestaciones), TODAS bajo el
# macroproceso 6 "Prestaciones Asociadas" según MAE_1. Default: 100% -> 601.
DEFAULT_PROCESO_GPA = 601
RECURSOS_GPA_REP3 = [6101, 6201, 6301, 6401, 6501, 6601]


def leer_tabla_actividad_st(file_bytes):
    """Lee una tabla ST_x para REP_3: además de CÓDIGO RECURSO / MONTO ANUAL
    ACTIVADO / TOTAL GASTO ANUAL (igual que para REP_2), detecta la columna
    CÓDIGO ACTIVIDAD (tolerante a variantes: 'CODIGO ACTIVIDAD',
    'CODIGO DE ACTIVIDAD', 'CÓDIGO DE ACTIVIDAD', etc. y errores de tipeo).
    Devuelve lista de (cod_recurso, monto_activado, gasto_no_activado, cod_actividad)
    y una lista de avisos (columna faltante, código de actividad o recurso
    no reconocido según MAE_1/MAE_2)."""
    avisos_tabla = []
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    header_row = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    idx_recurso = _encontrar_columna(header_row, [["CODIGO", "RECURSO"], ["COD", "RECURSO"]])
    idx_activado = _encontrar_columna(header_row, [["MONTO", "ANUAL", "ACTIVADO"], ["MONTO", "ACTIVADO"]])
    idx_gasto = _encontrar_columna(
        header_row,
        [["TOTAL", "GASTO", "ANUAL"], ["GASTO", "ANUAL", "NO", "ACTIVADO"], ["TOTAL", "GASTO"], ["GASTO"]],
        evitar=["ACTIVADO", "%"],
    )
    idx_actividad = _encontrar_columna(
        header_row,
        [["CODIGO", "ACTIVIDAD"], ["COD", "ACTIVIDAD"], ["CODIGO", "DE", "ACTIVIDAD"]],
    )

    faltantes = []
    if idx_recurso is None:
        faltantes.append("CÓDIGO RECURSO")
    if idx_activado is None:
        faltantes.append("MONTO ANUAL ACTIVADO")
    if idx_gasto is None:
        faltantes.append("TOTAL GASTO ANUAL")
    if idx_actividad is None:
        faltantes.append("CÓDIGO ACTIVIDAD")
    if faltantes:
        raise ValueError(f"No se identificaron las columnas: {', '.join(faltantes)}")

    filas = []
    recursos_invalidos = set()
    actividades_invalidas = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        cod_recurso = row[idx_recurso]
        cod_actividad = row[idx_actividad]
        if cod_recurso is None or cod_actividad is None:
            continue
        cod_recurso_num = _a_numero(cod_recurso)
        cod_recurso_num = int(cod_recurso_num) if cod_recurso_num == int(cod_recurso_num) else cod_recurso_num
        cod_actividad_num = int(_a_numero(cod_actividad))

        if cod_recurso_num not in RECURSO_NOMBRE:
            recursos_invalidos.add(cod_recurso_num)
        if cod_actividad_num not in CODIGOS_ACTIVIDAD_VALIDOS:
            actividades_invalidas.add(cod_actividad_num)

        filas.append((cod_recurso_num, _a_numero(row[idx_activado]), _a_numero(row[idx_gasto]), cod_actividad_num))

    if recursos_invalidos:
        avisos_tabla.append(
            f"CÓDIGO(S) RECURSO no reconocido(s) según MAE_2: {sorted(recursos_invalidos)}"
        )
    if actividades_invalidas:
        avisos_tabla.append(
            f"CÓDIGO(S) ACTIVIDAD no reconocido(s) según MAE_1: {sorted(actividades_invalidas)}"
        )
    return filas, avisos_tabla

def build_rep3(fb3, ggm_proceso_params, ogg_proceso_params, ggm_params, ogg_params, st_params, st_files_raw,
               mei_proceso_params=None, gpa_proceso_params=None, mei_params=None, gpa_files_raw=None):
    """
    Construye la tabla REP_3 (Costos y Gastos No Activados de Servicios
    Regulados - Proceso), abriendo el gasto regulado (familias 11 y 12)
    de REP_2 en Código de Proceso, para las familias GRH, GCP, GGV, GGI
    (con tablas de apertura por actividad dedicadas), GGM y OGG (con
    asignación por defecto/parametrizable, ya que no tienen tabla de
    apertura por actividad) y ST (que trae CÓDIGO ACTIVIDAD en la misma
    tabla, detectado de forma flexible/tolerante a variantes de nombre).

    fb3: dict con bytes de GRH_8, GRH_11, GRH_12, GCP_5, GCP_6, GGV_4, GGV_5,
         GGV_6, GGI_5, GGI_6, GGM_1..5, OGG_5 (todos opcionales -> None si
         no se cargó)
    st_files_raw: dict {nombre_tabla: file_bytes} de las tablas ST subidas
                  (se leen aquí con leer_tabla_actividad_st, no con
                  leer_tabla_st, ya que se necesita también CÓDIGO ACTIVIDAD)
    """
    avisos3 = []
    mei_proceso_params = mei_proceso_params or {}
    gpa_proceso_params = gpa_proceso_params or {}
    mei_params = mei_params or {}
    gpa_files_raw = gpa_files_raw or {}
    agg3 = defaultdict(float)  # (cod_recurso, familia_servicio, cod_proceso) -> gasto

    def safe_rows(clave, etiqueta):
        b = fb3.get(clave)
        if b is None:
            avisos3.append(f"⚠️ REP_3: falta **{etiqueta}** — esa familia queda excluida de REP_3.")
            return []
        try:
            return read_rows(b)
        except Exception as e:
            avisos3.append(f"⚠️ REP_3: no se pudo leer **{etiqueta}** ({e}).")
            return []

    REGULADOS = {1101, 1201, 1202, 1203, 1204, 1205}

    # ================= GRH =================
    grh8 = safe_rows("grh8", "GRH_8")
    grh11 = safe_rows("grh11", "GRH_11")
    grh12 = safe_rows("grh12", "GRH_12")

    if grh8 and grh11:
        shares_servicio_reg = defaultdict(list)
        for r in grh11:
            cod_reg, cod_noreg, pct = r[7], r[8], r[9]
            cod = cod_reg if cod_reg != -1 else cod_noreg
            if cod in REGULADOS:
                shares_servicio_reg[(r[4], r[5])].append((cod, pct))

        shares_proceso_grh = defaultdict(lambda: defaultdict(float))
        for r in grh12:
            persona, cargo, cod_actividad, pct_act = r[4], r[5], r[7], r[8]
            shares_proceso_grh[(persona, cargo)][cod_proceso_de_actividad(cod_actividad)] += pct_act

        personas_sin_grh12 = set()
        for r in grh8:
            empresa, periodo, anio, sector, persona, cargo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r
            servicios = shares_servicio_reg.get((persona, cargo), [])
            if not servicios:
                continue
            procesos = shares_proceso_grh.get((persona, cargo))
            if not procesos:
                personas_sin_grh12.add((persona, cargo))
                continue
            for cod_serv, pct_serv in servicios:
                fam = familia(cod_serv)
                gasto_rf = total_gasto * pct_serv
                for proceso, pct_proc in procesos.items():
                    agg3[(cod_recurso, fam, proceso)] += gasto_rf * pct_proc
        if personas_sin_grh12:
            avisos3.append(f"⚠️ GRH: {len(personas_sin_grh12)} persona(s) con dedicación regulada sin apertura en GRH_12; se excluyen de REP_3.")

    # ================= GCP =================
    gcp5 = safe_rows("gcp5", "GCP_5")
    gcp6 = safe_rows("gcp6", "GCP_6")

    if gcp5:
        shares_proceso_gcp = defaultdict(lambda: defaultdict(float))
        for r in gcp6:
            persona, cargo, cod_recurso_g6, cod_actividad, pct = r[4], r[5], r[6], r[8], r[9]
            shares_proceso_gcp[(persona, cargo, cod_recurso_g6)][cod_proceso_de_actividad(cod_actividad)] += pct

        faltan_gcp6 = 0
        for r in gcp5:
            empresa, periodo, anio, sector, persona, cargo, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r
            if cod_reg == -1 or cod_reg not in REGULADOS:
                continue
            fam = familia(cod_reg)
            procesos = shares_proceso_gcp.get((persona, cargo, cod_recurso))
            if not procesos:
                faltan_gcp6 += 1
                continue
            for proceso, pct_proc in procesos.items():
                agg3[(cod_recurso, fam, proceso)] += gasto_no_act * pct_proc
        if faltan_gcp6:
            avisos3.append(f"⚠️ GCP: {faltan_gcp6} combinación(es) persona-cargo-recurso reguladas sin apertura en GCP_6; se excluyen de REP_3.")

    # ================= GGV =================
    ggv4 = safe_rows("ggv4", "GGV_4")
    ggv5 = safe_rows("ggv5", "GGV_5")
    ggv6 = safe_rows("ggv6", "GGV_6")

    if ggv4 and ggv5:
        shares_servicio_ggv = defaultdict(list)
        for r in ggv5:
            empresa, periodo, anio, sector, id_activo, total_no_act, cod_reg, cod_noreg, pct = r[:9]
            if cod_reg in REGULADOS:
                shares_servicio_ggv[id_activo].append((cod_reg, pct))

        shares_proceso_ggv = defaultdict(lambda: defaultdict(float))
        for r in ggv6:
            id_activo, cod_actividad, pct = r[4], r[6], r[7]
            shares_proceso_ggv[id_activo][cod_proceso_de_actividad(cod_actividad)] += pct

        activos_sin_ggv6 = set()
        for r in ggv4:
            empresa, periodo, anio, sector, id_activo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r[:10]
            for cod_serv, pct_serv in shares_servicio_ggv.get(id_activo, []):
                fam = familia(cod_serv)
                gasto_rf = total_gasto * pct_serv
                procesos = shares_proceso_ggv.get(id_activo)
                if not procesos:
                    activos_sin_ggv6.add(id_activo)
                    continue
                for proceso, pct_proc in procesos.items():
                    agg3[(cod_recurso, fam, proceso)] += gasto_rf * pct_proc
        if activos_sin_ggv6:
            avisos3.append(f"⚠️ GGV: {len(activos_sin_ggv6)} activo(s) con dedicación regulada sin apertura en GGV_6; se excluyen de REP_3.")

    # ================= GGI =================
    ggi5 = safe_rows("ggi5", "GGI_5")
    ggi6 = safe_rows("ggi6", "GGI_6")

    if ggi5:
        shares_proceso_ggi = defaultdict(lambda: defaultdict(float))
        for r in ggi6:
            id_inmueble, cod_proceso_directo, pct = r[4], r[6], r[7]
            shares_proceso_ggi[id_inmueble][cod_proceso_directo] += pct

        inmuebles_sin_ggi6 = set()
        for r in ggi5:
            empresa, periodo, anio, sector, id_inmueble, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r[:10]
            if cod_reg == -1 or cod_reg not in REGULADOS:
                continue
            fam = familia(cod_reg)
            procesos = shares_proceso_ggi.get(id_inmueble)
            if not procesos:
                inmuebles_sin_ggi6.add(id_inmueble)
                continue
            for proceso, pct_proc in procesos.items():
                agg3[(cod_recurso, fam, proceso)] += gasto_no_act * pct_proc
        if inmuebles_sin_ggi6:
            avisos3.append(f"⚠️ GGI: {len(inmuebles_sin_ggi6)} inmueble(s) con dedicación regulada sin apertura en GGI_6; se excluyen de REP_3.")

    # ================= GGM (sin tabla de actividad: default/parametrizable) =================
    ggm_tablas = [safe_rows(f"ggm{i}", f"GGM_{i}") for i in range(1, 6)]
    if any(ggm_tablas):
        ggm_regulado_por_recurso = defaultdict(float)  # solo la porción a familia 11 (servicio 1101 u overrides de GGM_PARAMS)
        for tabla in ggm_tablas:
            for r in tabla:
                cod_recurso, total_gasto = r[4], r[-1]
                overrides = ggm_params.get(cod_recurso, [])
                pct_reg = 1.0 - sum(p for _, p in overrides)  # % que va a servicio 1101 (familia 11)
                ggm_regulado_por_recurso[cod_recurso] += total_gasto * pct_reg

        for cod_recurso, gasto in ggm_regulado_por_recurso.items():
            if gasto == 0:
                continue
            overrides_proceso = ggm_proceso_params.get(cod_recurso)
            if overrides_proceso:
                pct_default = 1.0 - sum(p for _, p in overrides_proceso)
                proceso_default = DEFAULT_PROCESO_GGM.get(cod_recurso, 504)
                agg3[(cod_recurso, 11, proceso_default)] += gasto * pct_default
                for proceso, pct in overrides_proceso:
                    agg3[(cod_recurso, 11, proceso)] += gasto * pct
            else:
                proceso_default = DEFAULT_PROCESO_GGM.get(cod_recurso, 504)
                agg3[(cod_recurso, 11, proceso_default)] += gasto

    # ================= OGG (sin tabla de actividad: default/parametrizable) =================
    ogg5 = safe_rows("ogg5", "OGG_5")
    if ogg5:
        ogg_regulado_por_recurso = defaultdict(float)
        for r in ogg5:
            cod_recurso, monto_act, total_gasto = r[4], r[9], r[11]
            overrides = ogg_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)
            ogg_regulado_por_recurso[cod_recurso] += total_gasto * pct_reg

        for cod_recurso, gasto in ogg_regulado_por_recurso.items():
            if gasto == 0:
                continue
            overrides_proceso = ogg_proceso_params.get(cod_recurso)
            proceso_default = DEFAULT_PROCESO_OGG.get(cod_recurso, DEFAULT_PROCESO_OGG_FALLBACK)
            if overrides_proceso:
                pct_default = 1.0 - sum(p for _, p in overrides_proceso)
                agg3[(cod_recurso, 11, proceso_default)] += gasto * pct_default
                for proceso, pct in overrides_proceso:
                    agg3[(cod_recurso, 11, proceso)] += gasto * pct
            else:
                agg3[(cod_recurso, 11, proceso_default)] += gasto

    # ================= ST (trae CÓDIGO ACTIVIDAD en la misma tabla) =================
    avisos_st_detalle = []
    for nombre_tabla, file_bytes in st_files_raw.items():
        try:
            filas_st, avisos_tabla = leer_tabla_actividad_st(file_bytes)
        except Exception as e:
            avisos3.append(f"⚠️ ST (REP_3): no se pudo leer **{nombre_tabla}** ({e}); se excluye.")
            continue
        for aviso in avisos_tabla:
            avisos_st_detalle.append(f"{nombre_tabla}: {aviso}")
        for cod_recurso, monto_act, total_gasto, cod_actividad in filas_st:
            overrides = st_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)  # % regulado (servicio 1101, familia 11)
            gasto_regulado_fila = total_gasto * pct_reg
            if gasto_regulado_fila == 0:
                continue
            proceso = cod_proceso_de_actividad(cod_actividad)
            agg3[(cod_recurso, 11, proceso)] += gasto_regulado_fila
    if avisos_st_detalle:
        avisos3.append(f"⚠️ ST (REP_3): validación de códigos — {'; '.join(avisos_st_detalle[:10])}" + (" ..." if len(avisos_st_detalle) > 10 else ""))

    # ================= MEI =================
    # MEI_1 (4101, Productos químicos): SIN columna de actividad -> default/parametrizable.
    # MEI_2 (4102), MEI_3 (4103), MEI_4 (4104-4106): SÍ traen CÓDIGO ACTIVIDAD
    # por fila (como ST) -> se deriva el proceso directamente.
    mei1 = safe_rows("mei1", "MEI_1")
    mei2 = safe_rows("mei2", "MEI_2")
    mei3 = safe_rows("mei3", "MEI_3")
    mei4 = safe_rows("mei4", "MEI_4")

    if mei1:
        mei1_regulado_por_recurso = defaultdict(float)
        for r in mei1:
            cod_recurso, total_gasto = r[6], r[22]
            overrides = mei_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)
            mei1_regulado_por_recurso[cod_recurso] += total_gasto * pct_reg
        for cod_recurso, gasto in mei1_regulado_por_recurso.items():
            if gasto == 0:
                continue
            overrides_proceso = mei_proceso_params.get(cod_recurso)
            if overrides_proceso:
                pct_default = 1.0 - sum(p for _, p in overrides_proceso)
                default_split = DEFAULT_PROCESO_MEI1.get(cod_recurso, [(504, 1.0)])
                for proceso, pct_d in default_split:
                    agg3[(cod_recurso, 11, proceso)] += gasto * pct_default * pct_d
                for proceso, pct in overrides_proceso:
                    agg3[(cod_recurso, 11, proceso)] += gasto * pct
            else:
                for proceso, pct_d in DEFAULT_PROCESO_MEI1.get(cod_recurso, [(504, 1.0)]):
                    agg3[(cod_recurso, 11, proceso)] += gasto * pct_d

    def procesar_mei_con_actividad(tabla, idx_recurso, idx_gasto, idx_actividad, etiqueta):
        if not tabla:
            return
        for r in tabla:
            cod_recurso, total_gasto, cod_actividad = r[idx_recurso], r[idx_gasto], r[idx_actividad]
            if cod_actividad is None:
                continue
            overrides = mei_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)
            gasto_regulado = total_gasto * pct_reg
            if gasto_regulado == 0:
                continue
            proceso = cod_proceso_de_actividad(cod_actividad)
            agg3[(cod_recurso, 11, proceso)] += gasto_regulado

    procesar_mei_con_actividad(mei2, 4, 16, 21, "MEI_2")
    procesar_mei_con_actividad(mei3, 4, 19, 9, "MEI_3")
    procesar_mei_con_actividad(mei4, 4, 16, 6, "MEI_4")

    # ================= GPA =================
    # Cada tabla GPA_x tiene 100% de su gasto en un servicio regulado FIJO
    # (1201-1205) y, por defecto, se asigna 100% al proceso 601 "Prestaciones
    # Asociadas" (parametrizable por código de recurso).
    tablas_gpa_faltantes = [t for t in GPA_TABLES if t not in gpa_files_raw]
    if tablas_gpa_faltantes and gpa_files_raw:
        avisos3.append(f"⚠️ REP_3-GPA: faltan {len(tablas_gpa_faltantes)} tabla(s): {', '.join(tablas_gpa_faltantes)}.")
    for nombre_tabla, file_bytes in gpa_files_raw.items():
        servicio_fijo = GPA_TABLE_TO_SERVICIO.get(nombre_tabla)
        if servicio_fijo is None:
            continue
        fam_gpa = familia(servicio_fijo)
        try:
            filas_gpa = leer_tabla_st(file_bytes)
        except Exception as e:
            avisos3.append(f"⚠️ REP_3-GPA: no se pudo leer {nombre_tabla} ({e}).")
            continue
        gpa_gasto_por_recurso = defaultdict(float)
        for cod_recurso, monto_act, total_gasto in filas_gpa:
            gpa_gasto_por_recurso[cod_recurso] += total_gasto
        for cod_recurso, gasto in gpa_gasto_por_recurso.items():
            if gasto == 0:
                continue
            overrides_proceso = gpa_proceso_params.get(cod_recurso)
            if overrides_proceso:
                pct_default = 1.0 - sum(p for _, p in overrides_proceso)
                agg3[(cod_recurso, fam_gpa, DEFAULT_PROCESO_GPA)] += gasto * pct_default
                for proceso, pct in overrides_proceso:
                    agg3[(cod_recurso, fam_gpa, proceso)] += gasto * pct
            else:
                agg3[(cod_recurso, fam_gpa, DEFAULT_PROCESO_GPA)] += gasto

    # ================= Normalizar (sumar EXACTAMENTE 100% por recurso+familia) =================
    grupos = defaultdict(list)
    for (cod_recurso, fam, proceso), gasto in agg3.items():
        if abs(gasto) < 1e-6:
            continue
        grupos[(cod_recurso, fam)].append((proceso, gasto))

    EMPRESA = PERIODO = ANIO = SECTOR = None
    for clave in ["grh8", "gcp5", "ggv4", "ggi5", "ggm1", "ogg5"]:
        b = fb3.get(clave)
        if b:
            filas_tmp = read_rows(b)
            if filas_tmp:
                EMPRESA, PERIODO, ANIO, SECTOR = filas_tmp[0][0], filas_tmp[0][1], filas_tmp[0][2], filas_tmp[0][3]
                break

    final_rows3 = []
    for (cod_recurso, fam), items in grupos.items():
        total_grupo = sum(g for _, g in items)
        fracciones = [g / total_grupo for _, g in items]
        pcts = redondear_para_sumar_100(fracciones, 4)
        for (proceso, gasto), pct in zip(items, pcts):
            final_rows3.append([EMPRESA, PERIODO, ANIO, SECTOR, cod_recurso, fam, proceso, round(pct * 100, 2), round(gasto, 2)])

    final_rows3.sort(key=lambda r: (r[4], r[5], r[6]))
    return final_rows3, avisos3



# ============================================================================
# CYG - Costos y Gastos por Recurso (CYG_1 a CYG_4, CYG_8, CYG_9)
# ============================================================================
# CYG_1: Recursos - Servicios Continuos AP/AS (servicios 1101, 1102)
# CYG_2: Recursos - Prestaciones Asociadas Reguladas (servicios 1201-1205)
# CYG_3: Recursos - Prestaciones Asociadas No Reguladas (servicios 2101-2117)
# CYG_4: Recursos - Servicios No Regulados (servicios 2201-2214)
# CYG_8: Recursos de Servicios Continuos - Actividades (solo gasto regulado)
# CYG_9: Recursos de Servicios No Regulados - Servicios Prestados a Terceros
#        (abre por ID CLIENTE usando MCO_42 + ING_4)
#
# IMPORTANTE: a diferencia de REP_2 (que colapsa los servicios en su FAMILIA
# 11/12/22) y de REP_3 (que colapsa las actividades en su PROCESO), las
# tablas CYG necesitan el código de SERVICIO ESPECÍFICO (ej. 1201, no solo
# "12") y el código de ACTIVIDAD ESPECÍFICO (ej. 4010101, no solo "401").
# Por eso CYG se construye SIEMPRE desde las tablas fuente, nunca desde
# REP_2/REP_3 ya agregados. REP_2 y REP_3 solo se usan como checkpoint de
# validación de cuadratura.

# --- Mapeo CÓDIGO OBRA TIPO NBI -> CÓDIGO ACTIVIDAD (exclusivo de MEI_1) ---
# Aportado por el usuario: reemplaza cualquier supuesto/estimación previa.
# El código de proceso se deriva de esta actividad (primeros 3 dígitos).
OBRA_NBI_A_ACTIVIDAD = {
    101: 1030401,  # Captación en Río (proceso 103)
    102: 1030402,  # Captación en Canal (proceso 103)
    103: 1030403,  # Captación en Lago o Embalse (proceso 103)
    104: 1030404,  # Captación en Mar (proceso 103)
    201: 1030302,  # Captación mediante Drenes y Galerías (proceso 103)
    202: 1030304,  # Captación mediante Punteras (proceso 103)
    203: 1030303,  # Captación mediante Sondajes (proceso 103)
    204: 1030301,  # Captación mediante Norias (proceso 103)
    301: 1031101,  # Plantas Elevadoras de Agua Potable Tipo A (proceso 103)
    302: 1031101,  # Plantas Elevadoras de Agua Potable Tipo B (proceso 103)
    303: 1031101,  # Plantas Elevadoras de Agua Potable Tipo C (proceso 103)
    304: 1031101,  # Plantas Elevadoras de Agua Potable Tipo D (proceso 103)
    305: 1031101,  # Plantas Elevadoras de Agua Potable Tipo E (proceso 103)
    351: 2020301,  # Plantas Elevadoras de Aguas Servidas (proceso 202)
    401: 1050101,  # Estanques Semienterrados y Enterrados (proceso 105)
    402: 1050102,  # Estanques Elevados (proceso 105)
    501: 1030601,  # Plantas de Tratamiento de Agua Potable excepto Osmosis Inversa (proceso 103)
    502: 1030604,  # Plantas de Tratamiento de Agua Potable de Osmosis Inversa (proceso 103)
    601: 1030801,  # Sistemas de Desinfección de Agua Potable (proceso 103)
    701: 1030902,  # Sistemas de Fluoración (proceso 103)
    801: 1060101,  # Red de Distribución (proceso 106)
    901: 2030301,  # Red de Recolección (proceso 203)
    1001: 3080101,  # Arranques (proceso 308)
    1002: 3070302,  # Medidores (proceso 307)
    1003: 3110301,  # Uniones Domiciliarias (proceso 311)
    1101: 1040101,  # Conducciones de AP (proceso 104)
    1151: 2030101,  # Conducciones de AS (proceso 203)
    1201: 2040404,  # Tabla General de Sistemas de Tratamiento de Aguas Servidas (proceso 204)
    1202: 2040501,  # Pretratamiento de Aguas Servidas (proceso 204)
    1203: 2040401,  # Tratamiento Primario de Aguas Servidas (proceso 204)
    1204: 2040401,  # Tratamiento Secundario de Aguas Servidas (proceso 204)
    1205: 2040404,  # Desinfección y Decloración en Plantas de Tratamiento de Aguas Servidas (proceso 204)
    1206: 2040404,  # Línea de Lodos de Plantas de Tratamiento de Aguas Servidas (proceso 204)
    1207: 2040601,  # Emisario Submarino - Información General (proceso 204)
    1208: 2040601,  # Emisario Submarino - Información por Tramos (proceso 204)
    1209: 2041001,  # Control Olores, Generador y Aforos en Plantas de Tratamiento de Aguas Servidas (proceso 204)
    1402: 3010901,  # Macromedidores (proceso 301)
    1403: 3011101,  # Reductoras de Presión (proceso 301)
    1404: 3011001,  # Anti Golpe de Ariete (proceso 301)
}

# --- Catálogo completo Proceso -> [Actividades] (para defaults GGI/GGM/OGG/GPA) ---
PROCESO_A_ACTIVIDADES = {
    101: [1010101, 1010201, 1010301, 1010302, 1010303, 1010304, 1010305, 1010306, 1010307],
    102: [1020101, 1020201, 1020301, 1020302, 1020303, 1020304],
    103: [1030101, 1030201, 1030202, 1030301, 1030302, 1030303, 1030304, 1030401, 1030402, 1030403, 1030404, 1030501, 1030502, 1030503, 1030504, 1030505, 1030601, 1030602, 1030603, 1030604, 1030605, 1030701, 1030702, 1030703, 1030704, 1030705, 1030801, 1030802, 1030803, 1030901, 1030902, 1030903, 1031001, 1031002, 1031101, 1031201, 1031202],
    104: [1040101, 1040201, 1040301, 1040302, 1040303, 1040304, 1040305, 1040306, 1040307, 1040308, 1040309, 1040401],
    105: [1050101, 1050102],
    106: [1060101, 1060102, 1060201, 1060202, 1060203, 1060204, 1060205, 1060301, 1060302],
    201: [2010101, 2010201, 2010301, 2010302, 2010303, 2010304, 2010305, 2010306, 2010307],
    202: [2020101, 2020102, 2020103, 2020201, 2020301],
    203: [2030101, 2030201, 2030301],
    204: [2040101, 2040102, 2040103, 2040104, 2040201, 2040301, 2040401, 2040402, 2040403, 2040404, 2040405, 2040501, 2040601, 2040701, 2040702, 2040703, 2040704, 2040705, 2040801, 2040802, 2040901, 2040902, 2040903, 2040904, 2041001],
    301: [3010101, 3010201, 3010301, 3010401, 3010501, 3010601, 3010701, 3010801, 3010901, 3011001, 3011101, 3011201, 3011301],
    302: [3020101, 3020201, 3020301, 3020401, 3020501],
    303: [3030101, 3030201, 3030301, 3030302, 3030303, 3030304, 3030401, 3030402, 3030403, 3030404, 3030501, 3030502, 3030503, 3030504, 3030505, 3030601, 3030602, 3030603, 3030701, 3030702, 3030703, 3030801, 3030802, 3030901, 3031001, 3031002],
    304: [3040101, 3040201, 3040301, 3040401, 3040501],
    305: [3050101, 3050201, 3050301, 3050302, 3050303, 3050401, 3050402, 3050501, 3050502],
    306: [3060101, 3060201, 3060301, 3060302, 3060303, 3060304, 3060401, 3060402, 3060403, 3060404, 3060501, 3060502, 3060503, 3060504],
    307: [3070101, 3070201, 3070301, 3070302],
    308: [3080101, 3080102, 3080103],
    309: [3090101, 3090201, 3090301, 3090302, 3090303, 3090401, 3090402, 3090403, 3090501, 3090502, 3090503, 3090504],
    310: [3100101, 3100201, 3100301, 3100302, 3100303, 3100304, 3100305, 3100401, 3100402, 3100403, 3100404, 3100405, 3100501, 3100502, 3100503, 3100504],
    311: [3110101, 3110201, 3110301, 3110302, 3110303, 3110304],
    312: [3120101, 3120201, 3120301, 3120401, 3120402, 3120403, 3120404, 3120405, 3120501, 3120601],
    313: [3130101, 3130201, 3130301, 3130401, 3130501],
    401: [4010101, 4010102, 4010103, 4010104, 4010105, 4010201, 4010202, 4010203, 4010204, 4010205, 4010206, 4010207, 4010301, 4010302, 4010303, 4010304, 4010401, 4010402, 4010403, 4010404, 4010405, 4010406, 4010407, 4010408],
    402: [4020101, 4020102, 4020103, 4020201, 4020202, 4020203, 4020301, 4020302, 4020303, 4020304, 4020305, 4020306, 4020307, 4020308, 4020309, 4020310, 4020311, 4020312, 4020313, 4020314, 4020315, 4020316, 4020317, 4020318, 4020319, 4020320, 4020321, 4020322, 4020323, 4020324, 4020325, 4020326, 4020327, 4020401, 4020402, 4020403, 4020501, 4020502, 4020503, 4020601, 4020602],
    403: [4030101, 4030102, 4030103, 4030104, 4030105, 4030106, 4030107, 4030108, 4030201, 4030202, 4030203, 4030204, 4030205, 4030206, 4030207, 4030208, 4030209, 4030210, 4030211, 4030301, 4030302, 4030303, 4030304, 4030305, 4030306, 4030307, 4030308, 4030309, 4030401, 4030402, 4030403, 4030404],
    501: [5010101, 5010102, 5010103, 5010104, 5010105, 5010106, 5010107, 5010108, 5010201, 5010202, 5010203, 5010204, 5010205, 5010206, 5010301, 5010302, 5010303, 5010304, 5010305, 5010306, 5010307, 5010308, 5010401, 5010402, 5010403, 5010404, 5010405, 5010406, 5010407, 5010501, 5010502, 5010503, 5010504, 5010505, 5010506, 5010601, 5010602, 5010603, 5010604, 5010605, 5010606, 5010607, 5010608],
    502: [5020101, 5020102, 5020103, 5020104, 5020105, 5020106, 5020107, 5020108, 5020109, 5020110, 5020111, 5020112, 5020113, 5020114, 5020115, 5020201, 5020202, 5020203, 5020204, 5020205, 5020206, 5020207, 5020208, 5020209, 5020210, 5020211, 5020212, 5020301, 5020302, 5020303, 5020304, 5020305, 5020306],
    503: [5030101, 5030102, 5030103, 5030104, 5030105, 5030106, 5030107, 5030108, 5030109, 5030110, 5030111, 5030112, 5030113, 5030114, 5030115, 5030116, 5030117, 5030118, 5030119, 5030120, 5030121, 5030122, 5030123, 5030124, 5030125, 5030201, 5030202, 5030203, 5030204, 5030205, 5030206, 5030207, 5030208, 5030209, 5030210, 5030211, 5030212, 5030213],
    504: [5040101, 5040102, 5040103, 5040104, 5040105, 5040106, 5040107, 5040201, 5040202, 5040203, 5040204, 5040205, 5040206, 5040207, 5040208, 5040209, 5040210, 5040301, 5040302, 5040303, 5040304],
    505: [5050101, 5050102, 5050103, 5050104, 5050105, 5050106, 5050107, 5050108, 5050109, 5050110, 5050111, 5050112, 5050113, 5050114, 5050115, 5050116],
    506: [5060101, 5060102, 5060103, 5060104, 5060105, 5060106, 5060107, 5060201, 5060202, 5060203, 5060204, 5060301, 5060302, 5060303, 5060304, 5060401, 5060402, 5060403, 5060404, 5060405, 5060406, 5060501, 5060502, 5060503, 5060504, 5060505, 5060506, 5060507, 5060508, 5060509],
    601: [6010101, 6010102, 6010103, 6010104, 6010105, 6010106, 6010107, 6010108, 6010109, 6010110, 6010111, 6010112, 6010113, 6010114, 6010201, 6010202, 6010203, 6010204, 6010301, 6010302, 6010303, 6010304, 6010305, 6010306, 6010401, 6010402, 6010403, 6010404, 6010501, 6010502, 6010503, 6010504, 6010505, 6010506, 6010507, 6010508, 6010509, 6010601, 6010602],
    701: [7010101, 7010102, 7010103, 7010104, 7010105, 7010201, 7010202, 7010203, 7010204, 7010205, 7010206, 7010207, 7010301, 7010302, 7010303, 7010304, 7010305],
    702: [7020101, 7020102, 7020103, 7020104, 7020105, 7020106],
    703: [7030101, 7030102, 7030103, 7030104, 7030105, 7030106, 7030201, 7030202, 7030203, 7030204, 7030301, 7030302, 7030303, 7030304],
    704: [7040101, 7040102, 7040103, 7040104, 7040105, 7040201, 7040202, 7040203, 7040204, 7040205],
    705: [7050101, 7050102, 7050103, 7050201, 7050202, 7050301, 7050302],
    706: [7060101],
}

def actividades_de_proceso(cod_proceso):
    return PROCESO_A_ACTIVIDADES.get(cod_proceso, [])

# --- Defaults de actividad para familias sin apertura (equitativo dentro del proceso) ---
# GGI (proceso 504, 100% de sus 21 actividades, sin excluir ninguna)
DEFAULT_ACTIVIDADES_GGI = {504: actividades_de_proceso(504)}

# GGM: 2401-2409 -> proceso 505 (16 act.); 2410-2411 -> proceso 504 (21 act., completas)
DEFAULT_ACTIVIDADES_GGM = {505: actividades_de_proceso(505), 504: actividades_de_proceso(504)}

# OGG: 2501-2502 -> proceso 501 (43 act.); resto -> proceso 502 (33 act.)
DEFAULT_ACTIVIDADES_OGG = {501: actividades_de_proceso(501), 502: actividades_de_proceso(502)}

# GPA: cada tabla reparte equitativamente SOLO entre las actividades de su
# propio subproceso (ej. GPA_1/6101 -> solo las 14 actividades de "Control
# Directo de Riles", no las 39 de todo el proceso 601).
def _acts_subproceso(prefijo5, catalogo_601):
    return [a for a in catalogo_601 if str(a)[:5] == prefijo5]

_acts_601 = actividades_de_proceso(601)
DEFAULT_ACTIVIDADES_GPA = {
    6101: _acts_subproceso("60101", _acts_601),
    6201: _acts_subproceso("60102", _acts_601),
    6301: _acts_subproceso("60103", _acts_601),
    6401: _acts_subproceso("60104", _acts_601),
    6501: _acts_subproceso("60105", _acts_601),
    6601: _acts_subproceso("60106", _acts_601),
}

def build_cyg_core(fb, ggm_params, ogg_params, mei_params, st_params, st_files_raw,
                    ggm_proceso_params, ogg_proceso_params, mei_proceso_params, gpa_proceso_params,
                    gpa_files_raw):
    """
    Construye, desde las tablas fuente (NO desde REP_2/REP_3), dos
    agregaciones con granularidad completa:

    agg_serv[(cod_servicio_especifico, cod_recurso)] = [gasto_no_act, monto_act]
        -> insumo para CYG_1, CYG_2, CYG_3, CYG_4 (agrupado por servicio
           específico, ej. 1201 separado de 1202).

    agg_act[(cod_actividad_especifica, cod_recurso)] = [gasto_no_act, monto_act]
        -> insumo para CYG_8 (SOLO la porción de gasto regulada, familias 11/12).

    Devuelve también 'avisos' con lo que falta o no se pudo determinar.
    """
    avisos = []
    agg_serv = defaultdict(lambda: [0.0, 0.0])
    agg_act = defaultdict(lambda: [0.0, 0.0])
    REGULADOS = {1101, 1201, 1202, 1203, 1204, 1205}

    def safe_rows(clave, etiqueta):
        b = fb.get(clave)
        if b is None:
            avisos.append(f"CYG: falta **{etiqueta}** — esa familia queda excluida.")
            return []
        try:
            return read_rows(b)
        except Exception as e:
            avisos.append(f"CYG: no se pudo leer **{etiqueta}** ({e}).")
            return []

    # ================= GRH =================
    grh8 = safe_rows("grh8", "GRH_8")
    grh11 = safe_rows("grh11", "GRH_11")
    grh12 = safe_rows("grh12", "GRH_12")

    shares_servicio_grh = defaultdict(list)  # (persona,cargo) -> [(servicio, pct), ...] TODOS los servicios
    for r in grh11:
        cod_reg, cod_noreg, pct = r[7], r[8], r[9]
        cod = cod_reg if cod_reg != -1 else cod_noreg
        shares_servicio_grh[(r[4], r[5])].append((cod, pct))

    shares_proceso_grh = defaultdict(lambda: defaultdict(float))  # (persona,cargo) -> {actividad: pct}
    # OJO: GRH_12 da % por ACTIVIDAD (no colapsado a proceso) -- se usa tal cual para CYG_8.
    shares_actividad_grh = defaultdict(lambda: defaultdict(float))
    for r in grh12:
        persona, cargo, cod_actividad, pct_act = r[4], r[5], r[7], r[8]
        shares_actividad_grh[(persona, cargo)][int(cod_actividad)] += pct_act

    personas_sin_grh12 = set()
    for r in grh8:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r
        for cod_serv, pct_serv in shares_servicio_grh.get((persona, cargo), []):
            gasto_serv = total_gasto * pct_serv
            monto_act_serv = monto_act * pct_serv
            agg_serv[(cod_serv, cod_recurso)][0] += gasto_serv
            agg_serv[(cod_serv, cod_recurso)][1] += monto_act_serv
            if cod_serv in REGULADOS:
                actividades = shares_actividad_grh.get((persona, cargo))
                if not actividades:
                    personas_sin_grh12.add((persona, cargo))
                    continue
                for cod_actividad, pct_act_i in actividades.items():
                    agg_act[(cod_actividad, cod_recurso)][0] += gasto_serv * pct_act_i
                    agg_act[(cod_actividad, cod_recurso)][1] += monto_act_serv * pct_act_i
    if personas_sin_grh12:
        avisos.append(f"CYG_8-GRH: {len(personas_sin_grh12)} persona(s) reguladas sin apertura en GRH_12.")

    # ================= GCP =================
    gcp5 = safe_rows("gcp5", "GCP_5")
    gcp6 = safe_rows("gcp6", "GCP_6")

    shares_actividad_gcp = defaultdict(lambda: defaultdict(float))  # (persona,cargo,recurso) -> {actividad: pct}
    for r in gcp6:
        persona, cargo, cod_recurso_g6, cod_actividad, pct = r[4], r[5], r[6], r[8], r[9]
        shares_actividad_gcp[(persona, cargo, cod_recurso_g6)][int(cod_actividad)] += pct

    faltan_gcp6 = 0
    for r in gcp5:
        empresa, periodo, anio, sector, persona, cargo, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        agg_serv[(cod_serv, cod_recurso)][0] += gasto_no_act
        # GCP_5 no trae desglose de monto activado por servicio -> se deja en 0 aquí
        if cod_reg in REGULADOS:
            actividades = shares_actividad_gcp.get((persona, cargo, cod_recurso))
            if not actividades:
                faltan_gcp6 += 1
                continue
            for cod_actividad, pct_a in actividades.items():
                agg_act[(cod_actividad, cod_recurso)][0] += gasto_no_act * pct_a
    if faltan_gcp6:
        avisos.append(f"CYG_8-GCP: {faltan_gcp6} combinación(es) reguladas sin apertura en GCP_6.")

    # ================= GGV =================
    ggv4 = safe_rows("ggv4", "GGV_4")
    ggv5 = safe_rows("ggv5", "GGV_5")
    ggv6 = safe_rows("ggv6", "GGV_6")

    shares_servicio_ggv = defaultdict(list)
    for r in ggv5:
        empresa, periodo, anio, sector, id_activo, total_no_act, cod_reg, cod_noreg, pct = r[:9]
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        shares_servicio_ggv[id_activo].append((cod_serv, pct))

    shares_actividad_ggv = defaultdict(lambda: defaultdict(float))
    for r in ggv6:
        id_activo, cod_actividad, pct = r[4], r[6], r[7]
        shares_actividad_ggv[id_activo][int(cod_actividad)] += pct

    activos_sin_ggv6 = set()
    for r in ggv4:
        empresa, periodo, anio, sector, id_activo, cod_recurso, monto_anual, monto_act, pct_act, total_gasto = r[:10]
        for cod_serv, pct_serv in shares_servicio_ggv.get(id_activo, []):
            gasto_serv = total_gasto * pct_serv
            monto_act_serv = monto_act * pct_serv
            agg_serv[(cod_serv, cod_recurso)][0] += gasto_serv
            agg_serv[(cod_serv, cod_recurso)][1] += monto_act_serv
            if cod_serv in REGULADOS:
                actividades = shares_actividad_ggv.get(id_activo)
                if not actividades:
                    activos_sin_ggv6.add(id_activo)
                    continue
                for cod_actividad, pct_a in actividades.items():
                    agg_act[(cod_actividad, cod_recurso)][0] += gasto_serv * pct_a
                    agg_act[(cod_actividad, cod_recurso)][1] += monto_act_serv * pct_a
    if activos_sin_ggv6:
        avisos.append(f"CYG_8-GGV: {len(activos_sin_ggv6)} activo(s) regulados sin apertura en GGV_6.")

    # ================= GGI =================
    ggi5 = safe_rows("ggi5", "GGI_5")
    ggi6 = safe_rows("ggi6", "GGI_6")

    shares_proceso_ggi = defaultdict(lambda: defaultdict(float))
    for r in ggi6:
        id_inmueble, cod_proceso_directo, pct = r[4], r[6], r[7]
        shares_proceso_ggi[id_inmueble][cod_proceso_directo] += pct

    for r in ggi5:
        empresa, periodo, anio, sector, id_inmueble, cod_recurso, gasto_no_act, cod_reg, cod_noreg, pct = r[:10]
        cod_serv = cod_reg if cod_reg != -1 else cod_noreg
        agg_serv[(cod_serv, cod_recurso)][0] += gasto_no_act
        if cod_reg in REGULADOS:
            # GGI no tiene tabla de actividad -> reparto equitativo por defecto
            # entre las actividades del proceso 504 (parametrizable)
            overrides = DEFAULT_ACTIVIDADES_GGI  # {proceso: [actividades]}; sin parametrización propia aún
            actividades_504 = overrides.get(504, [])
            if actividades_504:
                pct_igual = 1.0 / len(actividades_504)
                for cod_actividad in actividades_504:
                    agg_act[(cod_actividad, cod_recurso)][0] += gasto_no_act * pct_igual

    # ================= GGM =================
    ggm_tablas = [safe_rows(f"ggm{i}", f"GGM_{i}") for i in range(1, 6)]
    ggm_regulado_por_recurso = defaultdict(lambda: [0.0, 0.0])  # [gasto_no_act, monto_act]
    for tabla in ggm_tablas:
        for r in tabla:
            cod_recurso, monto_act, total_gasto = r[4], r[-3], r[-1]
            overrides = ggm_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)
            agg_serv[(1101, cod_recurso)][0] += total_gasto * pct_reg
            agg_serv[(1101, cod_recurso)][1] += monto_act  # monto activado SIEMPRE 100% regulado
            for cod_serv_over, pct_o in overrides:
                agg_serv[(cod_serv_over, cod_recurso)][0] += total_gasto * pct_o
                # el monto activado no se traspasa a servicios no regulados
            ggm_regulado_por_recurso[cod_recurso][0] += total_gasto * pct_reg
            ggm_regulado_por_recurso[cod_recurso][1] += monto_act

    def _repartir_actividades_igual(cod_recurso, gasto_g, gasto_a, proceso_default, overrides_proceso, defaults_map):
        if gasto_g == 0 and gasto_a == 0:
            return
        if overrides_proceso:
            pct_default = 1.0 - sum(p for _, p in overrides_proceso)
            acts_default = defaults_map.get(proceso_default) or actividades_de_proceso(proceso_default)
            if acts_default:
                pct_igual = pct_default / len(acts_default)
                for a in acts_default:
                    agg_act[(a, cod_recurso)][0] += gasto_g * pct_igual
                    agg_act[(a, cod_recurso)][1] += gasto_a * pct_igual
            for proc_ov, pct_ov in overrides_proceso:
                acts_ov = actividades_de_proceso(proc_ov)
                if acts_ov:
                    pct_igual_ov = pct_ov / len(acts_ov)
                    for a in acts_ov:
                        agg_act[(a, cod_recurso)][0] += gasto_g * pct_igual_ov
                        agg_act[(a, cod_recurso)][1] += gasto_a * pct_igual_ov
        else:
            acts_default = defaults_map.get(proceso_default) or actividades_de_proceso(proceso_default)
            if acts_default:
                pct_igual = 1.0 / len(acts_default)
                for a in acts_default:
                    agg_act[(a, cod_recurso)][0] += gasto_g * pct_igual
                    agg_act[(a, cod_recurso)][1] += gasto_a * pct_igual

    for cod_recurso, (gasto_g, gasto_a) in ggm_regulado_por_recurso.items():
        proceso_default = DEFAULT_PROCESO_GGM.get(cod_recurso, 504)
        overrides_proceso = ggm_proceso_params.get(cod_recurso)
        _repartir_actividades_igual(cod_recurso, gasto_g, gasto_a, proceso_default, overrides_proceso, DEFAULT_ACTIVIDADES_GGM)

    # ================= OGG =================
    ogg5 = safe_rows("ogg5", "OGG_5")
    ogg_regulado_por_recurso = defaultdict(lambda: [0.0, 0.0])
    for r in ogg5:
        cod_recurso, monto_act, total_gasto = r[4], r[9], r[11]
        overrides = ogg_params.get(cod_recurso, [])
        pct_reg = 1.0 - sum(p for _, p in overrides)
        agg_serv[(1101, cod_recurso)][0] += total_gasto * pct_reg
        agg_serv[(1101, cod_recurso)][1] += monto_act  # monto activado SIEMPRE 100% regulado
        for cod_serv_over, pct_o in overrides:
            agg_serv[(cod_serv_over, cod_recurso)][0] += total_gasto * pct_o
            # el monto activado no se traspasa a servicios no regulados
        ogg_regulado_por_recurso[cod_recurso][0] += total_gasto * pct_reg
        ogg_regulado_por_recurso[cod_recurso][1] += monto_act

    for cod_recurso, (gasto_g, gasto_a) in ogg_regulado_por_recurso.items():
        proceso_default = DEFAULT_PROCESO_OGG.get(cod_recurso, DEFAULT_PROCESO_OGG_FALLBACK)
        overrides_proceso = ogg_proceso_params.get(cod_recurso)
        _repartir_actividades_igual(cod_recurso, gasto_g, gasto_a, proceso_default, overrides_proceso, DEFAULT_ACTIVIDADES_OGG)

    # ================= MEI =================
    mei1 = safe_rows("mei1", "MEI_1")
    mei2 = safe_rows("mei2", "MEI_2")
    mei3 = safe_rows("mei3", "MEI_3")
    mei4 = safe_rows("mei4", "MEI_4")

    # --- MEI_1: mapeo EXACTO vía CÓDIGO OBRA TIPO NBI (no estimación) ---
    obras_no_mapeadas = set()
    for r in mei1:
        cod_recurso, cod_obra_nbi, monto_act, total_gasto = r[6], r[5], r[20], r[22]
        overrides = mei_params.get(cod_recurso, [])
        pct_reg = 1.0 - sum(p for _, p in overrides)
        gasto_reg = total_gasto * pct_reg
        act_reg = monto_act  # monto activado SIEMPRE 100% regulado
        agg_serv[(1101, cod_recurso)][0] += gasto_reg
        agg_serv[(1101, cod_recurso)][1] += act_reg
        for cod_serv_over, pct_o in overrides:
            agg_serv[(cod_serv_over, cod_recurso)][0] += total_gasto * pct_o
            # el monto activado no se traspasa a servicios no regulados
        if gasto_reg == 0 and act_reg == 0:
            continue
        cod_actividad = OBRA_NBI_A_ACTIVIDAD.get(cod_obra_nbi)
        if cod_actividad is None:
            obras_no_mapeadas.add(cod_obra_nbi)
            continue
        agg_act[(cod_actividad, cod_recurso)][0] += gasto_reg
        agg_act[(cod_actividad, cod_recurso)][1] += act_reg
    if obras_no_mapeadas:
        avisos.append(f"CYG_8-MEI_1: código(s) OBRA TIPO NBI sin mapeo conocido: {sorted(obras_no_mapeadas)}.")

    def procesar_mei_con_actividad(tabla, idx_recurso, idx_gasto, idx_activado, idx_actividad):
        for r in tabla:
            cod_recurso, total_gasto, monto_act, cod_actividad = r[idx_recurso], r[idx_gasto], r[idx_activado], r[idx_actividad]
            if cod_actividad is None:
                continue
            overrides = mei_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)
            gasto_reg = total_gasto * pct_reg
            act_reg = monto_act  # monto activado SIEMPRE 100% regulado
            agg_serv[(1101, cod_recurso)][0] += gasto_reg
            agg_serv[(1101, cod_recurso)][1] += act_reg
            for cod_serv_over, pct_o in overrides:
                agg_serv[(cod_serv_over, cod_recurso)][0] += total_gasto * pct_o
                # el monto activado no se traspasa a servicios no regulados
            if gasto_reg == 0 and act_reg == 0:
                continue
            agg_act[(int(cod_actividad), cod_recurso)][0] += gasto_reg
            agg_act[(int(cod_actividad), cod_recurso)][1] += act_reg

    procesar_mei_con_actividad(mei2, 4, 16, 17, 21)
    procesar_mei_con_actividad(mei3, 4, 19, 17, 9)
    procesar_mei_con_actividad(mei4, 4, 16, 14, 6)

    # ================= ST =================
    tablas_st_faltantes = [t for t in ST_TABLES if t not in st_files_raw]
    if tablas_st_faltantes and st_files_raw:
        avisos.append(f"CYG: faltan {len(tablas_st_faltantes)} tabla(s) ST: {', '.join(tablas_st_faltantes)}.")
    for nombre_tabla, file_bytes in st_files_raw.items():
        try:
            filas_st, avisos_tabla = leer_tabla_actividad_st(file_bytes)
        except Exception as e:
            avisos.append(f"CYG-ST: no se pudo leer {nombre_tabla} ({e}).")
            continue
        for a in avisos_tabla:
            avisos.append(f"CYG-ST {nombre_tabla}: {a}")
        for cod_recurso, monto_act, total_gasto, cod_actividad in filas_st:
            overrides = st_params.get(cod_recurso, [])
            pct_reg = 1.0 - sum(p for _, p in overrides)
            gasto_reg = total_gasto * pct_reg
            act_reg = monto_act  # monto activado SIEMPRE 100% regulado
            agg_serv[(1101, cod_recurso)][0] += gasto_reg
            agg_serv[(1101, cod_recurso)][1] += act_reg
            for cod_serv_over, pct_o in overrides:
                agg_serv[(cod_serv_over, cod_recurso)][0] += total_gasto * pct_o
                # el monto activado no se traspasa a servicios no regulados
            if gasto_reg == 0 and act_reg == 0:
                continue
            agg_act[(cod_actividad, cod_recurso)][0] += gasto_reg
            agg_act[(cod_actividad, cod_recurso)][1] += act_reg

    # ================= GPA =================
    tablas_gpa_faltantes = [t for t in GPA_TABLES if t not in gpa_files_raw]
    if tablas_gpa_faltantes and gpa_files_raw:
        avisos.append(f"CYG-GPA: faltan {len(tablas_gpa_faltantes)} tabla(s): {', '.join(tablas_gpa_faltantes)}.")
    for nombre_tabla, file_bytes in gpa_files_raw.items():
        servicio_fijo = GPA_TABLE_TO_SERVICIO.get(nombre_tabla)
        if servicio_fijo is None:
            continue
        try:
            filas_gpa = leer_tabla_st(file_bytes)
        except Exception as e:
            avisos.append(f"CYG-GPA: no se pudo leer {nombre_tabla} ({e}).")
            continue
        gpa_por_recurso = defaultdict(lambda: [0.0, 0.0])
        for cod_recurso, monto_act, total_gasto in filas_gpa:
            agg_serv[(servicio_fijo, cod_recurso)][0] += total_gasto
            agg_serv[(servicio_fijo, cod_recurso)][1] += monto_act
            gpa_por_recurso[cod_recurso][0] += total_gasto
            gpa_por_recurso[cod_recurso][1] += monto_act
        for cod_recurso, (gasto_g, gasto_a) in gpa_por_recurso.items():
            overrides_proceso = gpa_proceso_params.get(cod_recurso)
            acts_default_map = {DEFAULT_PROCESO_GPA: DEFAULT_ACTIVIDADES_GPA.get(cod_recurso, [])}
            _repartir_actividades_igual(cod_recurso, gasto_g, gasto_a, DEFAULT_PROCESO_GPA, overrides_proceso, acts_default_map)

    EMPRESA = PERIODO = ANIO = SECTOR = None
    for clave in ["grh8", "gcp5", "ggv4", "ggi5", "ggm1", "ogg5"]:
        b = fb.get(clave)
        if b:
            filas_tmp = read_rows(b)
            if filas_tmp:
                EMPRESA, PERIODO, ANIO, SECTOR = filas_tmp[0][0], filas_tmp[0][1], filas_tmp[0][2], filas_tmp[0][3]
                break

    return agg_serv, agg_act, avisos, (EMPRESA, PERIODO, ANIO, SECTOR)

def _normalizar_por_clave(items_por_clave, decimales=4):
    """items_por_clave: {clave_agrupadora: [(sub_clave, gasto), ...]}
    Devuelve {clave_agrupadora: [(sub_clave, pct, gasto), ...]} con pct
    sumando EXACTAMENTE 1.0 dentro de cada clave_agrupadora (método de mayor
    residuo)."""
    resultado = {}
    for clave, items in items_por_clave.items():
        total = sum(g for _, g in items)
        if total == 0:
            continue
        fracciones = [g / total for _, g in items]
        pcts = redondear_para_sumar_100(fracciones, decimales)
        resultado[clave] = [(sub, pct, g) for (sub, g), pct in zip(items, pcts)]
    return resultado


def build_cyg_1_a_4(agg_serv, empresa_periodo_anio_sector):
    """Arma CYG_1 (1101,1102), CYG_2 (1201-1205), CYG_3 (2101-2117),
    CYG_4 (2201-2214). Normaliza % por SERVICIO ESPECÍFICO (no por recurso)."""
    EMPRESA, PERIODO, ANIO, SECTOR = empresa_periodo_anio_sector
    rangos = {
        "CYG_1": set([1101, 1102]),
        "CYG_2": set([1201, 1202, 1203, 1204, 1205]),
        "CYG_3": set(range(2101, 2118)),
        "CYG_4": set(range(2201, 2215)),
    }
    resultados = {}
    for nombre_tabla, servicios_validos in rangos.items():
        # separar no-activado y activado (normalización independiente, igual que REP_2)
        items_no_act = defaultdict(list)  # servicio -> [(recurso, gasto), ...]
        items_act = defaultdict(list)
        for (cod_serv, cod_recurso), (g, a) in agg_serv.items():
            if cod_serv not in servicios_validos:
                continue
            if abs(g) > 1e-9:
                items_no_act[cod_serv].append((cod_recurso, g))
            if abs(a) > 1e-9:
                items_act[cod_serv].append((cod_recurso, a))

        norm_no_act = _normalizar_por_clave(items_no_act)
        norm_act = _normalizar_por_clave(items_act)

        filas = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])  # (servicio,recurso) -> [pct_no_act,monto_no_act,pct_act,monto_act]
        for servicio, lista in norm_no_act.items():
            for recurso, pct, g in lista:
                filas[(servicio, recurso)][0] = pct
                filas[(servicio, recurso)][1] = round(g, 2)
        for servicio, lista in norm_act.items():
            for recurso, pct, a in lista:
                filas[(servicio, recurso)][2] = pct
                filas[(servicio, recurso)][3] = round(a, 2)

        final_rows = []
        for (servicio, recurso), (pct_na, monto_na, pct_a, monto_a) in sorted(filas.items()):
            if monto_na == 0 and monto_a == 0:
                continue
            final_rows.append([EMPRESA, PERIODO, ANIO, SECTOR, servicio, recurso,
                                round(pct_na * 100, 2), monto_na, round(pct_a * 100, 2), monto_a])
        resultados[nombre_tabla] = final_rows
    return resultados


def build_cyg_8(agg_act, empresa_periodo_anio_sector):
    """CYG_8: (actividad, recurso, %no_act, monto_no_act, %act, monto_act).
    Normaliza % por RECURSO (sin distinguir familia de servicio)."""
    EMPRESA, PERIODO, ANIO, SECTOR = empresa_periodo_anio_sector

    items_no_act = defaultdict(list)  # recurso -> [(actividad, gasto), ...]
    items_act = defaultdict(list)
    for (cod_actividad, cod_recurso), (g, a) in agg_act.items():
        if abs(g) > 1e-9:
            items_no_act[cod_recurso].append((cod_actividad, g))
        if abs(a) > 1e-9:
            items_act[cod_recurso].append((cod_actividad, a))

    norm_no_act = _normalizar_por_clave(items_no_act)
    norm_act = _normalizar_por_clave(items_act)

    filas = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])  # (actividad,recurso) -> [pct_no_act,monto_no_act,pct_act,monto_act]
    for recurso, lista in norm_no_act.items():
        for actividad, pct, g in lista:
            filas[(actividad, recurso)][0] = pct
            filas[(actividad, recurso)][1] = round(g, 2)
    for recurso, lista in norm_act.items():
        for actividad, pct, a in lista:
            filas[(actividad, recurso)][2] = pct
            filas[(actividad, recurso)][3] = round(a, 2)

    final_rows = []
    for (actividad, recurso), (pct_na, monto_na, pct_a, monto_a) in sorted(filas.items()):
        if monto_na == 0 and monto_a == 0:
            continue
        final_rows.append([EMPRESA, PERIODO, ANIO, SECTOR, actividad, recurso,
                            round(pct_na * 100, 2), monto_na, round(pct_a * 100, 2), monto_a])
    return final_rows


PATRON_RUT_CHILENO = re.compile(r"^\d{6,9}-[0-9kK]$")


def es_rut_valido(id_cliente):
    """True si id_cliente tiene formato de RUT chileno válido (ej.
    '12345678-9' o '12345678-K'). Se usa para filtrar, en ING_4, las filas
    que corresponden a clientes reales (excluyendo provisiones, notas
    contables, referencias de facturación u otros textos que no son RUT)."""
    if id_cliente is None:
        return False
    return bool(PATRON_RUT_CHILENO.match(str(id_cliente).strip()))


def build_cyg_9(agg_serv, mco42_bytes, ing4_bytes, empresa_periodo_anio_sector):
    """CYG_9: abre servicios NO regulados (21xx, 22xx) por ID CLIENTE, usando
    MCO_42 (servicio -> id_ingreso) e ING_4 (id_ingreso -> id_cliente + monto
    anual de ingreso) para repartir proporcionalmente al ingreso de cada
    cliente. Solo se consideran clientes con ID CLIENTE en formato RUT
    chileno válido (ej. '12345678-9' o '12345678-K'); las demás filas de
    ING_4 (provisiones, notas contables, referencias de facturación, etc.)
    se EXCLUYEN del cálculo de proporción y se reportan en 'avisos' para que
    se corrijan en el origen. Si un servicio no está cubierto por MCO_42/
    ING_4 (o solo tiene entradas no-RUT), se declara con ID CLIENTE = '-1'
    (100%), según lo dispuesto en el diccionario SISS."""
    EMPRESA, PERIODO, ANIO, SECTOR = empresa_periodo_anio_sector
    avisos = []

    if mco42_bytes is None or ing4_bytes is None:
        avisos.append("CYG_9: falta MCO_42 y/o ING_4 — todos los clientes se declaran como '-1'.")
        mco = []
        ing = []
    else:
        mco = read_rows(mco42_bytes)
        ing = read_rows(ing4_bytes)

    servicio_a_ingreso = {r[4]: r[5] for r in mco}
    clientes_por_ingreso = defaultdict(list)
    no_rut_encontrados = []  # (id_ingreso, id_cliente, monto) para el aviso
    for r in ing:
        id_ingreso, id_cliente, monto = r[5], r[6], r[9]
        if es_rut_valido(id_cliente):
            clientes_por_ingreso[id_ingreso].append((id_cliente, monto))
        else:
            no_rut_encontrados.append((id_ingreso, id_cliente, monto))

    if no_rut_encontrados:
        monto_excluido = sum(m for _, _, m in no_rut_encontrados)
        ejemplos = ", ".join(f"'{c}'" for _, c, _ in no_rut_encontrados[:5])
        avisos.append(
            f"CYG_9: se excluyeron {len(no_rut_encontrados)} fila(s) de ING_4 con ID CLIENTE "
            f"sin formato RUT válido (monto total excluido de la base de reparto: "
            f"{monto_excluido:,.0f}). Revisar y corregir en el origen. Ejemplos: {ejemplos}"
            + (", ..." if len(no_rut_encontrados) > 5 else "") + "."
        )

    servicios_no_reg = set(range(2101, 2118)) | set(range(2201, 2215))

    filas = []
    for (cod_serv, cod_recurso), (g, a) in agg_serv.items():
        if cod_serv not in servicios_no_reg:
            continue
        if g == 0 and a == 0:
            continue

        id_ingreso = servicio_a_ingreso.get(cod_serv)
        clientes = clientes_por_ingreso.get(id_ingreso) if id_ingreso is not None else None

        if not clientes:
            # sin cobertura MCO_42/ING_4 -> 100% a ID CLIENTE "-1"
            filas.append([cod_serv, cod_recurso, "-1", g, a])
            continue

        total_ingreso = sum(c[1] for c in clientes)
        if total_ingreso == 0:
            filas.append([cod_serv, cod_recurso, "-1", g, a])
            continue

        # Reparto proporcional SIN redondear aún (el redondeo final, con piso
        # mínimo, se aplica una sola vez más abajo al consolidar por recurso;
        # redondear aquí prematuramente a 4 decimales podía dejar en 0.0000 —
        # y por lo tanto eliminar— a clientes muy pequeños relativos al total
        # de ingreso, antes de que el piso mínimo alcanzara a rescatarlos).
        for id_cliente, monto in clientes:
            pct_cliente = monto / total_ingreso
            filas.append([cod_serv, cod_recurso, id_cliente, g * pct_cliente, a * pct_cliente])

    # Consolidar filas duplicadas (mismo servicio,recurso,cliente pudiera repetirse si
    # el mismo cliente aparece en más de un ID INGRESO -- no debería, pero por robustez)
    consolidado = defaultdict(lambda: [0.0, 0.0])
    for cod_serv, cod_recurso, id_cliente, g, a in filas:
        consolidado[(cod_serv, cod_recurso, id_cliente)][0] += g
        consolidado[(cod_serv, cod_recurso, id_cliente)][1] += a

    # Normalizar % por RECURSO (no por servicio+recurso): el % representa la
    # porción del RECURSO asignada a cada cliente, combinando todos los
    # servicios no regulados que usan ese recurso.
    items_no_act = defaultdict(list)
    items_act = defaultdict(list)
    for (cod_serv, cod_recurso, id_cliente), (g, a) in consolidado.items():
        if abs(g) > 1e-9:
            items_no_act[cod_recurso].append(((cod_serv, id_cliente), g))
        if abs(a) > 1e-9:
            items_act[cod_recurso].append(((cod_serv, id_cliente), a))

    norm_no_act, fusion_no_act = normalizar_con_piso_minimo(items_no_act, piso=0.0001)
    norm_act, fusion_act = normalizar_con_piso_minimo(items_act, piso=0.0001)

    recursos_fusionados = {rec for rec, f in fusion_no_act.items() if f} | {rec for rec, f in fusion_act.items() if f}
    if recursos_fusionados:
        avisos.append(
            f"CYG_9: {len(recursos_fusionados)} recurso(s) con demasiados clientes para "
            f"que todos alcancen el piso mínimo de 0,01% individualmente; los clientes de "
            f"menor monto se agruparon bajo ID CLIENTE '-1'. Recursos afectados: "
            f"{sorted(recursos_fusionados)}."
        )

    filas_finales = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for cod_recurso, lista in norm_no_act.items():
        for (cod_serv, id_cliente), pct, g in lista:
            filas_finales[(cod_serv, cod_recurso, id_cliente)][0] = pct
            filas_finales[(cod_serv, cod_recurso, id_cliente)][1] = round(g, 2)
    for cod_recurso, lista in norm_act.items():
        for (cod_serv, id_cliente), pct, a in lista:
            filas_finales[(cod_serv, cod_recurso, id_cliente)][2] = pct
            filas_finales[(cod_serv, cod_recurso, id_cliente)][3] = round(a, 2)

    final_rows = []
    for (cod_serv, cod_recurso, id_cliente), (pct_na, monto_na, pct_a, monto_a) in sorted(
            filas_finales.items(), key=lambda x: (x[0][0], x[0][1], str(x[0][2]))):
        if monto_na == 0 and monto_a == 0:
            continue
        final_rows.append([EMPRESA, PERIODO, ANIO, SECTOR, id_cliente, cod_serv, cod_recurso,
                            round(pct_na * 100, 2), monto_na, round(pct_a * 100, 2), monto_a])

    return final_rows, avisos


# ---------------------------------------------------------------- UI --------
st.title("Generador REP_2 · Costos y Gastos por Familia de Servicios")
st.caption(
    "Consolida las familias GRH, GCP, GGV, GGI, GGM, OGG y MEI en la tabla "
    "REP_2 exigida por la SISS."
)

with st.sidebar:
    st.header("Archivos de entrada")
    st.caption("Todos los archivos son opcionales: puedes generar REP_2 con lo que tengas disponible.")
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
    st.markdown("**ST — Servicios Tercerizados (29 tablas)**")
    st.caption("Sube cualquier subconjunto de ST_3.xlsx a ST_34.xlsx (se identifican por nombre de archivo).")
    f_st_files = st.file_uploader(
        "Tablas ST (selección múltiple)", type="xlsx", key="st_files", accept_multiple_files=True
    )
    st.markdown("**GPA — Gasto Prestaciones Asociadas**")
    st.caption("100% del gasto va al servicio regulado fijo de cada tabla (no requiere parametrización).")
    f_gpa_files = st.file_uploader(
        "Tablas GPA (selección múltiple, GPA_1 a GPA_6)", type="xlsx", key="gpa_files", accept_multiple_files=True
    )
    st.markdown("**Plantilla (opcional)**")
    f_template = st.file_uploader("REP_2.xlsx (diccionario)", type="xlsx", key="template")

    archivos_regulares = [f_grh8, f_grh11, f_gcp4, f_gcp5, f_ggv4, f_ggv5, f_ggi5,
                           f_ggm1, f_ggm2, f_ggm3, f_ggm4, f_ggm5, f_ogg5,
                           f_mei1, f_mei2, f_mei3, f_mei4]
    hay_algo_cargado = any(archivos_regulares) or bool(f_st_files) or bool(f_gpa_files)

st.markdown(
    """
**Lógica aplicada**
- **GRH / GGV**: % de dedicación (por persona o por activo) aplicado a los montos de gasto.
- **GCP / GGI**: ya vienen abiertas por recurso y servicio.
- **GGM / OGG / MEI / ST**: sin apertura por servicio — por defecto 100% va al
  servicio regulado 1101; se puede parametrizar % a servicios no regulados
  en los paneles de abajo.
- **GPA**: 6 tablas, cada una con un servicio regulado FIJO (GPA_1→1201,
  GPA_2→1202, GPA_3→1203, GPA_4→1204, GPA_5→1205, GPA_6→1201). No requiere
  parametrización de servicios no regulados.
- **Generación parcial**: puedes generar REP_2 con solo algunas tablas
  cargadas. Las familias/recursos sin datos quedan en 0, y se muestra un
  aviso detallado de qué faltó o llegó vacío.
    """
)

st.subheader("Parametrización opcional — familias sin apertura por servicio")
st.caption("Por defecto todo el gasto se asigna 100% al servicio regulado 1101.")

with st.expander("Configurar parametrización GGM / OGG / MEI / ST", expanded=False):
    ggm_params = panel_parametrizacion("GGM (recursos 2401-2411)", RECURSOS_GGM, "ggm_overrides")
    st.divider()
    ogg_params = panel_parametrizacion("OGG (recursos 2501-2550)", RECURSOS_OGG, "ogg_overrides")
    st.divider()
    mei_params = panel_parametrizacion("MEI (recursos 4101-4106)", RECURSOS_MEI, "mei_overrides")
    st.divider()
    st.caption("ST: algunos códigos son compartidos por varias tablas (ej. 5110 en ST_22 a ST_30); la parametrización se aplica por código de recurso, no por tabla.")
    st_params = panel_parametrizacion("ST (Servicios Tercerizados)", RECURSOS_ST, "st_overrides")

if "ggm_overrides" not in st.session_state:
    ggm_params = {}
if "ogg_overrides" not in st.session_state:
    ogg_params = {}
if "mei_overrides" not in st.session_state:
    mei_params = {}
if "st_overrides" not in st.session_state:
    st_params = {}

# Validar overflow > 100%
def check_overflow(params):
    return {r: sum(p for _, p in lst) for r, lst in params.items() if sum(p for _, p in lst) > 1.0}

overflow = {}
overflow.update(check_overflow(ggm_params))
overflow.update(check_overflow(ogg_params))
overflow.update(check_overflow(mei_params))
overflow.update(check_overflow(st_params))
if overflow:
    st.error(f"La suma de % parametrizados supera 100% para el(los) recurso(s): {list(overflow.keys())}. Ajusta los valores.")

run = st.button("Generar REP_2", type="primary", disabled=not hay_algo_cargado or bool(overflow))

if not hay_algo_cargado:
    st.info("Sube al menos un archivo en el panel izquierdo para habilitar la generación (puede ser parcial).")

if run:
    fb = {
        "grh8": f_grh8.getvalue() if f_grh8 else None,
        "grh11": f_grh11.getvalue() if f_grh11 else None,
        "gcp4": f_gcp4.getvalue() if f_gcp4 else None,
        "gcp5": f_gcp5.getvalue() if f_gcp5 else None,
        "ggv4": f_ggv4.getvalue() if f_ggv4 else None,
        "ggv5": f_ggv5.getvalue() if f_ggv5 else None,
        "ggi5": f_ggi5.getvalue() if f_ggi5 else None,
        "ggm1": f_ggm1.getvalue() if f_ggm1 else None,
        "ggm2": f_ggm2.getvalue() if f_ggm2 else None,
        "ggm3": f_ggm3.getvalue() if f_ggm3 else None,
        "ggm4": f_ggm4.getvalue() if f_ggm4 else None,
        "ggm5": f_ggm5.getvalue() if f_ggm5 else None,
        "ogg5": f_ogg5.getvalue() if f_ogg5 else None,
        "mei1": f_mei1.getvalue() if f_mei1 else None,
        "mei2": f_mei2.getvalue() if f_mei2 else None,
        "mei3": f_mei3.getvalue() if f_mei3 else None,
        "mei4": f_mei4.getvalue() if f_mei4 else None,
    }

    st_files = {}
    avisos_carga_archivos = []
    for f in (f_st_files or []):
        tabla = identificar_tabla_st(f.name)
        if tabla is None:
            avisos_carga_archivos.append(f"⚠️ El archivo **{f.name}** no coincide con ninguna tabla ST_3..ST_34 conocida; se ignora.")
            continue
        try:
            st_files[tabla] = leer_tabla_st(f.getvalue())
        except Exception as e:
            avisos_carga_archivos.append(f"⚠️ No se pudo leer **{f.name}** ({tabla}): {e}. Se excluye del cálculo.")

    gpa_files = {}
    for f in (f_gpa_files or []):
        tabla = identificar_tabla_gpa(f.name)
        if tabla is None:
            avisos_carga_archivos.append(f"⚠️ El archivo **{f.name}** no coincide con ninguna tabla GPA_1..GPA_6 conocida; se ignora.")
            continue
        try:
            gpa_files[tabla] = leer_tabla_st(f.getvalue())  # misma lectura robusta por encabezado
        except Exception as e:
            avisos_carga_archivos.append(f"⚠️ No se pudo leer **{f.name}** ({tabla}): {e}. Se excluye del cálculo.")

    try:
        final_rows, checks, familia_map, by_recurso_planas, gpa_detalle, avisos = build_rep2(
            fb, ggm_params, ogg_params, mei_params, st_files, st_params, gpa_files
        )
        avisos = avisos_carga_archivos + avisos
    except Exception as e:
        st.error(f"Error procesando los archivos: {e}")
        st.stop()

    df = pd.DataFrame(final_rows, columns=HEADERS)
    if len(df) == 0:
        st.error("No se generó ninguna fila. Revisa que al menos una tabla tenga datos válidos.")
        st.stop()
    st.success(f"REP_2 generado con {len(df)} filas.")
    st.session_state['last_final_rows'] = final_rows
    st.session_state['last_familia_map'] = familia_map
    st.session_state['last_by_recurso_planas'] = by_recurso_planas
    st.session_state['last_gpa_detalle'] = gpa_detalle
    st.session_state['last_params_by_familia'] = {"GGM": ggm_params, "OGG": ogg_params, "MEI": mei_params, "ST": st_params}
    st.session_state['last_avisos'] = avisos

    if avisos:
        with st.expander(f"⚠️ Avisos de carga ({len(avisos)})", expanded=True):
            for aviso in avisos:
                st.markdown(f"- {aviso}")

    st.subheader("Validación de cuadratura por familia de gasto")
    st.caption("Solo se muestran las familias con datos cargados.")
    cols = st.columns(max(len(checks), 1))
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
            "% NO ACTIVADO ASIGNADO FAMILIA SERVICIOS": "{:.2f}%",
            "% ACTIVADO ASIGNADO FAMILIA SERVICIOS": "{:.2f}%",
            "GASTO ANUAL": "{:,.0f}",
            "MONTO ACTIVADO": "{:,.0f}",
        }),
        use_container_width=True,
    )

    params_by_familia = {"GGM": ggm_params, "OGG": ogg_params, "MEI": mei_params, "ST": st_params}
    excel_bytes = build_excel(
        final_rows, familia_map, by_recurso_planas, params_by_familia,
        gpa_detalle=gpa_detalle,
        avisos=avisos,
        template_bytes=f_template.getvalue() if f_template else None,
    )
    st.download_button(
        "Descargar REP_2.xlsx",
        data=excel_bytes,
        file_name="REP_2.xlsx",

        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ============================================================================
# PANEL DE VISUALIZACIÓN: evolución histórica y detección de anomalías
# ============================================================================
st.divider()
st.header("📊 Panel de Visualización")
st.caption(
    "Compara el año actual (generado arriba) contra un libro histórico "
    "consolidado (ej. REP_2_2020-2024.xlsx) para ver la evolución del gasto "
    "por Familia de Gasto, Código de Recurso y Familia de Servicio, además "
    "de detectar variaciones interanuales atípicas."
)

f_historico = st.file_uploader(
    "Libro histórico consolidado REP_2 (opcional, ej. REP_2_2020-2024.xlsx)",
    type="xlsx", key="historico_viz"
)


def cargar_historico(file_bytes):
    if file_bytes is None:
        return []
    wb_h = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws_h = wb_h.active
    return [r for r in ws_h.iter_rows(min_row=2, values_only=True) if r[0] is not None]


def gasto_total_fila(r):
    # Solo GASTO ANUAL (no activado). No se suma MONTO ACTIVADO.
    return r[7] or 0


def detectar_anomalias(evol_dict, nombre_dim):
    anomalias = []
    for key, serie_dict in evol_dict.items():
        anios_serie = sorted(serie_dict.keys())
        if len(anios_serie) < 2:
            continue
        valores = [serie_dict[a] for a in anios_serie]
        cambios = []
        for i in range(1, len(valores)):
            prev, curr = valores[i - 1], valores[i]
            pct_cambio = None if prev == 0 else (curr - prev) / prev
            cambios.append((anios_serie[i], prev, curr, pct_cambio))
        pct_validos = [c[3] for c in cambios if c[3] is not None]
        if len(pct_validos) >= 3:
            media = statistics.mean(pct_validos)
            desv = statistics.stdev(pct_validos)
        else:
            media, desv = None, None
        for anio_actual, prev, curr, pct_cambio in cambios:
            if pct_cambio is None:
                continue
            es_anomalia = False
            motivo = ""
            if desv and desv > 1e-9:
                z = (pct_cambio - media) / desv
                if abs(z) > 2:
                    es_anomalia = True
                    motivo = f"Variación atípica (z={z:.1f}) vs. patrón histórico"
            else:
                if abs(pct_cambio) > 0.5:
                    es_anomalia = True
                    motivo = "Variación mayor a 50% (serie corta)"
            if es_anomalia:
                anomalias.append({
                    "Dimensión": nombre_dim, "Grupo": key, "Año": anio_actual,
                    "Valor Año Anterior": round(prev, 0), "Valor Año Actual": round(curr, 0),
                    "% Variación": round(pct_cambio, 4), "Motivo": motivo,
                })
    return anomalias


final_rows_actual = st.session_state.get('last_final_rows', [])
historico_rows = cargar_historico(f_historico.getvalue() if f_historico else None)
combinado = list(historico_rows) + list(final_rows_actual)

if not combinado:
    st.info("Genera el REP_2 del año actual arriba y/o sube el libro histórico para ver el panel de visualización.")
else:
    anios_disponibles = sorted(set(r[2] for r in combinado))

    evol_familia = defaultdict(lambda: defaultdict(float))
    evol_recurso = defaultdict(lambda: defaultdict(float))
    evol_famserv = defaultdict(lambda: defaultdict(float))
    FAMSERV_NOMBRE = {11: '11 - Servicios Sanitarios (Agua/Alcantarillado)', 12: '12 - Servicios Sanitarios (Otros)', 22: '22 - Servicios No Regulados'}

    for r in combinado:
        fam = familia_de_recurso(r[4])
        evol_familia[fam][r[2]] += gasto_total_fila(r)
        evol_recurso[r[4]][r[2]] += gasto_total_fila(r)
        evol_famserv[r[5]][r[2]] += gasto_total_fila(r)

    nombre_recurso_map = {cod: f"{cod} - {nombre_recurso(cod)}" for cod in evol_recurso.keys()}
    famserv_map = {fs: FAMSERV_NOMBRE.get(fs, str(fs)) for fs in evol_famserv.keys()}

    vista = st.radio(
        "Ver evolución por:",
        ["Familia de Gasto", "Código de Recurso", "Familia de Servicio"],
        horizontal=True,
    )

    if vista == "Familia de Gasto":
        evol_dict, label_map = evol_familia, {k: k for k in evol_familia}
    elif vista == "Código de Recurso":
        evol_dict, label_map = evol_recurso, nombre_recurso_map
    else:
        evol_dict, label_map = evol_famserv, famserv_map

    tabla = pd.DataFrame({
        label_map[k]: [evol_dict[k].get(a, 0) for a in anios_disponibles]
        for k in evol_dict
    }, index=anios_disponibles)
    tabla.index.name = "Año"

    if vista == "Código de Recurso" and len(tabla.columns) > 15:
        top_cols = tabla.sum(axis=0).sort_values(ascending=False).head(15).index
        st.caption("Mostrando los 15 códigos de recurso con mayor gasto acumulado (usa la tabla completa para ver el resto).")
        st.line_chart(tabla[top_cols])
    else:
        st.line_chart(tabla)

    with st.expander("Ver tabla de datos"):
        st.dataframe(tabla.style.format("{:,.0f}"), use_container_width=True)

    # --- Detección de anomalías ---
    st.subheader("⚠️ Anomalías detectadas (variaciones interanuales atípicas)")
    anomalias_familia = detectar_anomalias(evol_familia, "Familia de Gasto")
    anomalias_recurso = detectar_anomalias({nombre_recurso_map[k]: v for k, v in evol_recurso.items()}, "Código de Recurso")
    anomalias_famserv = detectar_anomalias({famserv_map[k]: v for k, v in evol_famserv.items()}, "Familia de Servicio")
    todas_anomalias = anomalias_familia + anomalias_recurso + anomalias_famserv

    if todas_anomalias:
        df_anom = pd.DataFrame(todas_anomalias).sort_values("% Variación", key=lambda s: s.abs(), ascending=False)
        st.dataframe(
            df_anom.style.format({
                "Valor Año Anterior": "{:,.0f}",
                "Valor Año Actual": "{:,.0f}",
                "% Variación": "{:+.1%}",
            }).background_gradient(subset=["% Variación"], cmap="RdYlGn_r"),
            use_container_width=True,
        )
    else:
        st.success("No se detectaron variaciones interanuales atípicas con los criterios actuales.")

    st.caption(
        "Criterio: para series con ≥3 variaciones interanuales se usa z-score "
        "(|z|>2) sobre el propio histórico de cada grupo; para series más "
        "cortas se marca si la variación supera 50%."
    )

    # --- Descarga del Excel completo (REP_2 + Panel de Visualización) ------
    st.divider()
    if 'last_final_rows' not in st.session_state:
        st.info("Genera el REP_2 del año actual arriba para poder descargar el Excel con el panel de visualización incluido.")
    else:
        if st.button("📥 Preparar Excel con Panel de Visualización (REP_2 + evolución histórica)"):
            excel_base = build_excel(
                st.session_state['last_final_rows'],
                st.session_state['last_familia_map'],
                st.session_state['last_by_recurso_planas'],
                st.session_state['last_params_by_familia'],
                gpa_detalle=st.session_state.get('last_gpa_detalle'),
                avisos=st.session_state.get('last_avisos'),
                template_bytes=f_template.getvalue() if 'f_template' in dir() and f_template else None,
            )
            wb_completo = openpyxl.load_workbook(excel_base)

            header_fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
            header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
            data_font = Font(name="Arial", size=10)
            thin = Side(style="thin", color="D9D9D9")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            ws5 = wb_completo.create_sheet("Evolucion_Familia_Gasto")
            ws5.append(["Familia de Gasto"] + [str(a) for a in anios_disponibles])
            for c in range(1, len(anios_disponibles) + 2):
                cell = ws5.cell(row=1, column=c)
                cell.font = header_font; cell.fill = header_fill; cell.border = border
                cell.alignment = Alignment(horizontal="center")
            for fam_k in sorted(evol_familia.keys()):
                ws5.append([fam_k] + [round(evol_familia[fam_k].get(a, 0), 2) for a in anios_disponibles])
            for r in range(2, ws5.max_row + 1):
                for c in range(1, len(anios_disponibles) + 2):
                    cell = ws5.cell(row=r, column=c)
                    cell.font = data_font; cell.border = border
                    if c > 1:
                        cell.number_format = '#,##0;(#,##0);"-"'
            ws5.column_dimensions["A"].width = 40
            for c in range(2, len(anios_disponibles) + 2):
                ws5.column_dimensions[get_column_letter(c)].width = 16
            if ws5.max_row > 1 and len(anios_disponibles) > 1:
                chart1 = LineChart()
                chart1.title = "Evolución del Gasto por Familia"
                chart1.y_axis.title = "Gasto Anual ($)"
                chart1.x_axis.title = "Año"
                data_ref = Reference(ws5, min_col=2, max_col=len(anios_disponibles) + 1, min_row=1, max_row=ws5.max_row)
                cats_ref = Reference(ws5, min_col=1, min_row=2, max_row=ws5.max_row)
                chart1.add_data(data_ref, titles_from_data=True)
                chart1.set_categories(cats_ref)
                chart1.width = 26; chart1.height = 12
                ws5.add_chart(chart1, f"A{ws5.max_row + 3}")

            ws6 = wb_completo.create_sheet("Evolucion_Recurso")
            ws6.append(["Código Recurso", "Nombre Recurso"] + [str(a) for a in anios_disponibles])
            for c in range(1, len(anios_disponibles) + 3):
                cell = ws6.cell(row=1, column=c)
                cell.font = header_font; cell.fill = header_fill; cell.border = border
                cell.alignment = Alignment(horizontal="center")
            for cod in sorted(evol_recurso.keys()):
                ws6.append([cod, nombre_recurso(cod)] + [round(evol_recurso[cod].get(a, 0), 2) for a in anios_disponibles])
            for r in range(2, ws6.max_row + 1):
                for c in range(1, len(anios_disponibles) + 3):
                    cell = ws6.cell(row=r, column=c)
                    cell.font = data_font; cell.border = border
                    if c > 2:
                        cell.number_format = '#,##0;(#,##0);"-"'
            ws6.column_dimensions["A"].width = 14
            ws6.column_dimensions["B"].width = 45
            for c in range(3, len(anios_disponibles) + 3):
                ws6.column_dimensions[get_column_letter(c)].width = 16
            ws6.freeze_panes = "C2"
            if ws6.max_row > 1:
                ultima_col = len(anios_disponibles) + 2
                rango = f"{get_column_letter(ultima_col)}2:{get_column_letter(ultima_col)}{ws6.max_row}"
                ws6.conditional_formatting.add(rango, ColorScaleRule(
                    start_type="min", start_color="63BE7B",
                    mid_type="percentile", mid_value=50, mid_color="FFEB84",
                    end_type="max", end_color="F8696B"))

            ws7 = wb_completo.create_sheet("Evolucion_Familia_Servicio")
            ws7.append(["Familia de Servicio"] + [str(a) for a in anios_disponibles])
            for c in range(1, len(anios_disponibles) + 2):
                cell = ws7.cell(row=1, column=c)
                cell.font = header_font; cell.fill = header_fill; cell.border = border
                cell.alignment = Alignment(horizontal="center")
            for fs in sorted(evol_famserv.keys()):
                ws7.append([famserv_map[fs]] + [round(evol_famserv[fs].get(a, 0), 2) for a in anios_disponibles])
            for r in range(2, ws7.max_row + 1):
                for c in range(1, len(anios_disponibles) + 2):
                    cell = ws7.cell(row=r, column=c)
                    cell.font = data_font; cell.border = border
                    if c > 1:
                        cell.number_format = '#,##0;(#,##0);"-"'
            ws7.column_dimensions["A"].width = 42
            for c in range(2, len(anios_disponibles) + 2):
                ws7.column_dimensions[get_column_letter(c)].width = 16
            if ws7.max_row > 1 and len(anios_disponibles) > 1:
                chart2 = BarChart()
                chart2.type = "col"; chart2.grouping = "clustered"
                chart2.title = "Evolución del Gasto por Familia de Servicio"
                chart2.y_axis.title = "Gasto Anual ($)"
                chart2.x_axis.title = "Año"
                data2 = Reference(ws7, min_col=2, max_col=len(anios_disponibles) + 1, min_row=1, max_row=ws7.max_row)
                cats2 = Reference(ws7, min_col=1, min_row=2, max_row=ws7.max_row)
                chart2.add_data(data2, titles_from_data=True)
                chart2.set_categories(cats2)
                chart2.width = 24; chart2.height = 11
                ws7.add_chart(chart2, f"A{ws7.max_row + 3}")

            ws8 = wb_completo.create_sheet("Anomalias_Detectadas")
            ws8.append(["Dimensión", "Grupo", "Año con Anomalía", "Valor Año Anterior", "Valor Año Actual", "% Variación", "Motivo"])
            for c in range(1, 8):
                cell = ws8.cell(row=1, column=c)
                cell.font = header_font; cell.fill = header_fill; cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws8.row_dimensions[1].height = 30
            if todas_anomalias:
                for a_row in todas_anomalias:
                    ws8.append([a_row["Dimensión"], a_row["Grupo"], a_row["Año"], a_row["Valor Año Anterior"],
                                a_row["Valor Año Actual"], a_row["% Variación"], a_row["Motivo"]])
            else:
                ws8.append(["(Sin anomalías detectadas con los criterios actuales)", "", "", "", "", "", ""])
            for r in range(2, ws8.max_row + 1):
                for c in range(1, 8):
                    cell = ws8.cell(row=r, column=c)
                    cell.font = data_font; cell.border = border
                    if c in (4, 5):
                        cell.number_format = '#,##0;(#,##0);"-"'
                    if c == 6 and isinstance(ws8.cell(row=r, column=6).value, float):
                        cell.number_format = "+0.00%;-0.00%"
                        val = ws8.cell(row=r, column=6).value
                        cell.fill = PatternFill("solid", start_color="FFC7CE" if val > 0 else "C6EFCE",
                                                 end_color="FFC7CE" if val > 0 else "C6EFCE")
            ws8.column_dimensions["A"].width = 16
            ws8.column_dimensions["B"].width = 42
            ws8.column_dimensions["C"].width = 14
            ws8.column_dimensions["D"].width = 18
            ws8.column_dimensions["E"].width = 18
            ws8.column_dimensions["F"].width = 12
            ws8.column_dimensions["G"].width = 55
            ws8.freeze_panes = "A2"

            out_completo = io.BytesIO()
            wb_completo.save(out_completo)
            out_completo.seek(0)
            st.session_state['excel_completo_bytes'] = out_completo.getvalue()
            st.success("Excel listo. Usa el botón de descarga que apareció abajo.")

        if 'excel_completo_bytes' in st.session_state:
            st.download_button(
                "Descargar REP_2 + Panel de Visualización.xlsx",
                data=st.session_state['excel_completo_bytes'],
                file_name="REP_2_con_evolucion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ============================================================================
# GENERADOR REP_3 - Costos y Gastos No Activados de Servicios Regulados - Proceso
# ============================================================================
st.divider()
st.header("📋 Generar REP_3 (apertura por Código de Proceso)")
st.caption(
    "REP_3 toma solo el gasto REGULADO de REP_2 (familias 11 y 12) y lo abre "
    "por Código de Proceso. El Código de Proceso corresponde a los primeros "
    "3 dígitos del Código Actividad, según MAE_1 del Maestro SISS."
)

HEADERS_REP3 = [
    "CÓDIGO EMPRESA", "PERÍODO INFORMACIÓN", "AÑO INFORMADO",
    "CÓDIGO SECTOR DECRETO TARIFARIO", "CÓDIGO RECURSO",
    "CÓDIGO FAMILIA SERVICIOS REGULADOS", "CÓDIGO PROCESO",
    "% ASIGNADO PROCESO", "GASTO ANUAL",
]


def build_excel_rep3(final_rows3):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "REP_3"

    header_fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    data_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(HEADERS_REP3)
    for c in range(1, len(HEADERS_REP3) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 30

    for row in final_rows3:
        ws.append(row)

    for r in range(2, ws.max_row + 1):
        for c in range(1, len(HEADERS_REP3) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
            if c == 8:
                cell.number_format = "0.00"
            elif c == 9:
                cell.number_format = '#,##0;(#,##0);"-"'
            else:
                cell.alignment = Alignment(horizontal="center")

    widths = [14, 16, 14, 20, 14, 24, 14, 16, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # Hoja de detalle por proceso (nombre descriptivo)
    ws2 = wb.create_sheet("Detalle_Procesos")
    ws2.append(["Código Proceso", "Nombre Proceso", "Gasto Anual Total"])
    for c in range(1, 4):
        cell = ws2.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    por_proceso = defaultdict(float)
    for row in final_rows3:
        por_proceso[row[6]] += row[8]
    for proceso in sorted(por_proceso.keys()):
        ws2.append([proceso, nombre_proceso(proceso), round(por_proceso[proceso], 2)])
    for r in range(2, ws2.max_row + 1):
        for c in range(1, 4):
            cell = ws2.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
            if c == 3:
                cell.number_format = '#,##0;(#,##0);"-"'
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 55
    ws2.column_dimensions["C"].width = 20

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def panel_parametrizacion_proceso_generico(nombre, recursos_disponibles, session_key, defaults_map):
    if session_key not in st.session_state:
        st.session_state[session_key] = []
    st.markdown(f"**{nombre}**")
    col1, col2, col3, col4 = st.columns([1.2, 2.5, 1, 0.8])
    with col1:
        sel_recurso = st.selectbox("Código Recurso", recursos_disponibles, key=f"sel_recurso_proc_{session_key}")
    with col2:
        sel_proceso = st.selectbox(
            "Proceso destino",
            options=sorted(PROCESO_NOMBRE.keys()),
            format_func=lambda c: f"{c} - {PROCESO_NOMBRE[c]}",
            key=f"sel_proceso_{session_key}",
        )
    with col3:
        sel_pct = st.number_input(
            "% asignado", min_value=0.0, max_value=100.0, value=10.0,
            step=0.0000000001, format="%.10f", key=f"sel_pct_proc_{session_key}",
            help="Puedes ingresar hasta 10 decimales para minimizar la diferencia de redondeo en el monto resultante.",
        )
    with col4:
        st.write("")
        st.write("")
        if st.button("Agregar", key=f"add_proc_{session_key}"):
            st.session_state[session_key].append((sel_recurso, sel_proceso, sel_pct / 100.0))
    if st.session_state[session_key]:
        for i, (cod_r, cod_p, pct) in enumerate(st.session_state[session_key]):
            c1, c2, c3, c4 = st.columns([1.2, 2.5, 1, 0.8])
            c1.write(cod_r)
            c2.write(f"{cod_p} - {PROCESO_NOMBRE.get(cod_p, '')}")
            c3.write(f"{pct:.9%}")
            if c4.button("Quitar", key=f"del_proc_{session_key}_{i}"):
                st.session_state[session_key].pop(i)
                st.rerun()
    else:
        default_desc = ", ".join(f"{r}→{p}" for r, p in list(defaults_map.items())[:3])
        st.caption(f"Sin parametrización: se usa el default (ej. {default_desc}...).")
    params = defaultdict(list)
    for cod_r, cod_p, pct in st.session_state[session_key]:
        params[cod_r].append((cod_p, pct))
    return dict(params)


with st.expander("📂 Archivos y parametrización para REP_3", expanded=False):
    st.caption(
        "REP_3 reutiliza GRH_8/GRH_11, GCP_5, GGV_4/GGV_5, GGI_5 y las tablas ST "
        "ya cargadas arriba para REP_2. Aquí solo necesitas subir las tablas de "
        "apertura por actividad/proceso adicionales."
    )
    col1, col2 = st.columns(2)
    with col1:
        f_grh12 = st.file_uploader("GRH_12.xlsx", type="xlsx", key="grh12")
        f_gcp6 = st.file_uploader("GCP_6.xlsx", type="xlsx", key="gcp6")
    with col2:
        f_ggv6 = st.file_uploader("GGV_6.xlsx", type="xlsx", key="ggv6")
        f_ggi6 = st.file_uploader("GGI_6.xlsx", type="xlsx", key="ggi6")

    st.markdown("**Parametrización de proceso — GGM y OGG (sin tabla de apertura)**")
    st.caption(
        "Por defecto: GGM → Informática (505) para recursos de TI/telecom, "
        "Abastecimiento y Servicios Generales (504) para materiales; "
        "OGG → Dirección Superior (501) para gastos de Directorio, "
        "Administración y Finanzas (502) para el resto. Editable abajo."
    )
    ggm_proceso_params = panel_parametrizacion_proceso_generico(
        "GGM (recursos 2401-2411)", RECURSOS_GGM, "ggm_proceso_overrides", DEFAULT_PROCESO_GGM
    )
    st.divider()
    ogg_proceso_params = panel_parametrizacion_proceso_generico(
        "OGG (recursos 2501-2550)", RECURSOS_OGG, "ogg_proceso_overrides", DEFAULT_PROCESO_OGG
    )
    st.divider()
    st.caption(
        "MEI_1 (Productos químicos, sin actividad): por defecto 70% al proceso "
        "103 (tratamiento AP) y 30% al proceso 204 (tratamiento AS), en base a "
        "la distribución real de MEI_2 (Energía Eléctrica) entre esos mismos "
        "procesos de tratamiento. MEI_2/3/4 ya traen su propio Código Actividad "
        "y no requieren parametrización."
    )
    mei_proceso_params = panel_parametrizacion_proceso_generico(
        "MEI_1 (recurso 4101)", [4101], "mei_proceso_overrides", DEFAULT_PROCESO_MEI1
    )
    st.divider()
    st.caption(
        "GPA: por defecto 100% al proceso 601 'Prestaciones Asociadas' (todas "
        "las tablas GPA corresponden a esa categoría según MAE_1)."
    )
    gpa_proceso_params = panel_parametrizacion_proceso_generico(
        "GPA (recursos 6101, 6201, 6301, 6401, 6501, 6601)", RECURSOS_GPA_REP3, "gpa_proceso_overrides", {r: 601 for r in [6101, 6201, 6301, 6401, 6501, 6601]}
    )

    st.markdown("**Tablas GPA para REP_3 (mismas que subiste para REP_2)**")
    f_gpa_files_rep3 = st.file_uploader(
        "Tablas GPA (selección múltiple, GPA_1 a GPA_6)", type="xlsx", key="gpa_files_rep3", accept_multiple_files=True
    )

run_rep3 = st.button("Generar REP_3", type="primary")

if run_rep3:
    fb3 = {
        "grh8": f_grh8.getvalue() if f_grh8 else None,
        "grh11": f_grh11.getvalue() if f_grh11 else None,
        "grh12": f_grh12.getvalue() if f_grh12 else None,
        "gcp5": f_gcp5.getvalue() if f_gcp5 else None,
        "gcp6": f_gcp6.getvalue() if f_gcp6 else None,
        "ggv4": f_ggv4.getvalue() if f_ggv4 else None,
        "ggv5": f_ggv5.getvalue() if f_ggv5 else None,
        "ggv6": f_ggv6.getvalue() if f_ggv6 else None,
        "ggi5": f_ggi5.getvalue() if f_ggi5 else None,
        "ggi6": f_ggi6.getvalue() if f_ggi6 else None,
        "ggm1": f_ggm1.getvalue() if f_ggm1 else None,
        "ggm2": f_ggm2.getvalue() if f_ggm2 else None,
        "ggm3": f_ggm3.getvalue() if f_ggm3 else None,
        "ggm4": f_ggm4.getvalue() if f_ggm4 else None,
        "ggm5": f_ggm5.getvalue() if f_ggm5 else None,
        "ogg5": f_ogg5.getvalue() if f_ogg5 else None,
        "mei1": f_mei1.getvalue() if f_mei1 else None,
        "mei2": f_mei2.getvalue() if f_mei2 else None,
        "mei3": f_mei3.getvalue() if f_mei3 else None,
        "mei4": f_mei4.getvalue() if f_mei4 else None,
    }
    st_files_raw = {}
    for f in (f_st_files or []):
        tabla = identificar_tabla_st(f.name)
        if tabla:
            st_files_raw[tabla] = f.getvalue()

    gpa_files_raw_rep3 = {}
    for f in (f_gpa_files_rep3 or []):
        tabla = identificar_tabla_gpa(f.name)
        if tabla:
            gpa_files_raw_rep3[tabla] = f.getvalue()

    try:
        final_rows3, avisos3 = build_rep3(
            fb3, ggm_proceso_params, ogg_proceso_params, ggm_params, ogg_params, st_params, st_files_raw,
            mei_proceso_params=mei_proceso_params, gpa_proceso_params=gpa_proceso_params,
            mei_params=mei_params, gpa_files_raw=gpa_files_raw_rep3,
        )
    except Exception as e:
        st.error(f"Error generando REP_3: {e}")
        st.stop()

    if not final_rows3:
        st.error("No se generó ninguna fila de REP_3. Verifica que hayas cargado al menos GRH_8+GRH_11+GRH_12, o alguna otra familia con su tabla de proceso.")
    else:
        st.success(f"REP_3 generado con {len(final_rows3)} filas.")
        if avisos3:
            with st.expander(f"⚠️ Avisos REP_3 ({len(avisos3)})", expanded=True):
                for a in avisos3:
                    st.markdown(f"- {a}")

        df3 = pd.DataFrame(final_rows3, columns=HEADERS_REP3)
        st.dataframe(
            df3.style.format({"% ASIGNADO PROCESO": "{:.2f}%", "GASTO ANUAL": "{:,.0f}"}),
            use_container_width=True,
        )

        excel_rep3 = build_excel_rep3(final_rows3)
        st.download_button(
            "Descargar REP_3.xlsx",
            data=excel_rep3,
            file_name="REP_3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ============================================================================
# GENERADOR CYG - Costos y Gastos por Recurso (CYG_1-4, CYG_8, CYG_9)
# ============================================================================
st.divider()
st.header("📊 Generar tablas CYG")
st.caption(
    "Las tablas CYG (Costos y Gastos por Recurso) se construyen SIEMPRE desde "
    "las tablas fuente (no desde REP_2/REP_3, que ya colapsan el servicio "
    "específico y la actividad específica en familia/proceso). REP_2 y REP_3 "
    "se usan aquí solo como checkpoint de validación de cuadratura."
)
st.caption(
    "CYG_1: servicios 1101/1102 · CYG_2: servicios 1201-1205 · "
    "CYG_3: servicios no regulados 2101-2117 · CYG_4: servicios no regulados "
    "2201-2214 · CYG_8: apertura por actividad (solo gasto regulado) · "
    "CYG_9: servicios no regulados abiertos por ID Cliente (vía MCO_42 + ING_4)."
)

HEADERS_CYG_SERV = [
    "CÓDIGO EMPRESA", "PERÍODO INFORMACIÓN", "AÑO INFORMADO", "CÓDIGO SECTOR DECRETO TARIFARIO",
    "CÓDIGO SERVICIO", "CÓDIGO RECURSO",
    "% NO ACTIVADO ASIGNADO", "MONTO NO ACTIVADO", "% ACTIVADO ASIGNADO", "MONTO ACTIVADO",
]
HEADERS_CYG_8 = [
    "CÓDIGO EMPRESA", "PERÍODO INFORMACIÓN", "AÑO INFORMADO", "CÓDIGO SECTOR DECRETO TARIFARIO",
    "CÓDIGO ACTIVIDAD", "CÓDIGO RECURSO",
    "% NO ACTIVADO ASIGNADO ACTIVIDAD", "MONTO NO ACTIVADO", "% ACTIVADO ASIGNADO ACTIVIDAD", "MONTO ACTIVADO",
]
HEADERS_CYG_9 = [
    "CÓDIGO EMPRESA", "PERÍODO INFORMACIÓN", "AÑO INFORMADO", "CÓDIGO SECTOR DECRETO TARIFARIO",
    "ID CLIENTE", "CÓDIGO SERVICIO NO REGULADO", "CÓDIGO RECURSO",
    "% NO ACTIVADO ASIGNADO A CLIENTE", "MONTO NO ACTIVADO", "% ACTIVADO ASIGNADO A CLIENTE", "MONTO ACTIVADO",
]


def _agregar_hoja_cyg(wb, nombre_hoja, headers, filas, pct_cols, monto_cols):
    ws = wb.create_sheet(nombre_hoja)
    header_fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    data_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 30

    for row in filas:
        ws.append(row)

    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = data_font
            cell.border = border
            if c in pct_cols:
                cell.number_format = "0.00"
            elif c in monto_cols:
                cell.number_format = '#,##0;(#,##0);"-"'
            else:
                cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    return ws


def build_excel_cyg(cyg14, cyg8, cyg9, avisos_cyg):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for nombre in ["CYG_1", "CYG_2", "CYG_3", "CYG_4"]:
        ws = _agregar_hoja_cyg(wb, nombre, HEADERS_CYG_SERV, cyg14[nombre], pct_cols={7, 9}, monto_cols={8, 10})
        widths = [14, 16, 14, 20, 14, 14, 16, 18, 16, 18]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    ws8 = _agregar_hoja_cyg(wb, "CYG_8", HEADERS_CYG_8, cyg8, pct_cols={7, 9}, monto_cols={8, 10})
    widths8 = [14, 16, 14, 20, 14, 14, 20, 18, 20, 18]
    for i, w in enumerate(widths8, start=1):
        ws8.column_dimensions[get_column_letter(i)].width = w

    ws9 = _agregar_hoja_cyg(wb, "CYG_9", HEADERS_CYG_9, cyg9, pct_cols={8, 10}, monto_cols={9, 11})
    widths9 = [14, 16, 14, 20, 16, 20, 14, 18, 18, 18, 18]
    for i, w in enumerate(widths9, start=1):
        ws9.column_dimensions[get_column_letter(i)].width = w

    if avisos_cyg:
        wsa = wb.create_sheet("Avisos_CYG")
        wsa.append(["Aviso"])
        cell = wsa.cell(row=1, column=1)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        for a in avisos_cyg:
            wsa.append([a.replace("**", "")])
        for r in range(2, wsa.max_row + 1):
            c = wsa.cell(row=r, column=1)
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        wsa.column_dimensions["A"].width = 110

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


with st.expander("📂 Archivos adicionales para CYG_9 (MCO_42 + ING_4)", expanded=False):
    st.caption(
        "MCO_42 asigna cada servicio no regulado a un ID Ingreso. ING_4 abre "
        "cada ID Ingreso en ID Cliente + monto anual de ingreso. El gasto se "
        "reparte entre clientes proporcionalmente a su ingreso. Si un servicio "
        "no está cubierto, se declara con ID CLIENTE = '-1' (100%)."
    )
    f_mco42 = st.file_uploader("MCO_42.xlsx", type="xlsx", key="mco42")
    f_ing4 = st.file_uploader("ING_4.xlsx", type="xlsx", key="ing4")

run_cyg = st.button("Generar tablas CYG", type="primary")

if run_cyg:
    fb_cyg = {
        "grh8": f_grh8.getvalue() if f_grh8 else None,
        "grh11": f_grh11.getvalue() if f_grh11 else None,
        "grh12": f_grh12.getvalue() if f_grh12 else None,
        "gcp5": f_gcp5.getvalue() if f_gcp5 else None,
        "gcp6": f_gcp6.getvalue() if f_gcp6 else None,
        "ggv4": f_ggv4.getvalue() if f_ggv4 else None,
        "ggv5": f_ggv5.getvalue() if f_ggv5 else None,
        "ggv6": f_ggv6.getvalue() if f_ggv6 else None,
        "ggi5": f_ggi5.getvalue() if f_ggi5 else None,
        "ggi6": f_ggi6.getvalue() if f_ggi6 else None,
        "ggm1": f_ggm1.getvalue() if f_ggm1 else None,
        "ggm2": f_ggm2.getvalue() if f_ggm2 else None,
        "ggm3": f_ggm3.getvalue() if f_ggm3 else None,
        "ggm4": f_ggm4.getvalue() if f_ggm4 else None,
        "ggm5": f_ggm5.getvalue() if f_ggm5 else None,
        "ogg5": f_ogg5.getvalue() if f_ogg5 else None,
        "mei1": f_mei1.getvalue() if f_mei1 else None,
        "mei2": f_mei2.getvalue() if f_mei2 else None,
        "mei3": f_mei3.getvalue() if f_mei3 else None,
        "mei4": f_mei4.getvalue() if f_mei4 else None,
    }
    st_files_raw_cyg = {}
    for f in (f_st_files or []):
        tabla = identificar_tabla_st(f.name)
        if tabla:
            st_files_raw_cyg[tabla] = f.getvalue()
    gpa_files_raw_cyg = {}
    for f in (f_gpa_files_rep3 or []):
        tabla = identificar_tabla_gpa(f.name)
        if tabla:
            gpa_files_raw_cyg[tabla] = f.getvalue()

    try:
        agg_serv, agg_act, avisos_cyg, epas = build_cyg_core(
            fb_cyg, ggm_params, ogg_params, mei_params, st_params, st_files_raw_cyg,
            ggm_proceso_params, ogg_proceso_params, mei_proceso_params, gpa_proceso_params,
            gpa_files_raw_cyg,
        )
    except Exception as e:
        st.error(f"Error generando CYG: {e}")
        st.stop()

    cyg14 = build_cyg_1_a_4(agg_serv, epas)
    cyg8 = build_cyg_8(agg_act, epas)
    cyg9, avisos9 = build_cyg_9(
        agg_serv,
        f_mco42.getvalue() if f_mco42 else None,
        f_ing4.getvalue() if f_ing4 else None,
        epas,
    )
    avisos_cyg = avisos_cyg + avisos9

    total_filas = sum(len(v) for v in cyg14.values()) + len(cyg8) + len(cyg9)
    st.success(f"Tablas CYG generadas: {total_filas} filas en total.")
    for nombre, filas in cyg14.items():
        st.write(f"**{nombre}**: {len(filas)} filas")
    st.write(f"**CYG_8**: {len(cyg8)} filas")
    st.write(f"**CYG_9**: {len(cyg9)} filas")

    if avisos_cyg:
        with st.expander(f"⚠️ Avisos CYG ({len(avisos_cyg)})", expanded=True):
            for a in avisos_cyg:
                st.markdown(f"- {a}")

    excel_cyg = build_excel_cyg(cyg14, cyg8, cyg9, avisos_cyg)
    st.download_button(
        "Descargar CYG.xlsx",
        data=excel_cyg,
        file_name="CYG.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
