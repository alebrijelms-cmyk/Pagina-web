"""
Proyecto 5 - Preparación de datos (REPDA -> Excel de 15,000 registros x 10 campos)

Entrada : ReporteRepda.csv (REPDA/CONAGUA, ~540,000 títulos, 14 columnas)
Salida  : Pozos_Riego_15000.xlsx  (hojas: Datos, Registro_Ejecucion, Mapeo_Campos)

Campos REALES (derivados del CSV): id_titulo_concesion, volumen_extraido_m3, fecha_concesion,
                                   nombre_titular, estado, uso_principal
Campos SIMULADOS (el REPDA no los publica): municipio, profundidad_pozo_m, estatus_pozo,
                                            caudal_litros_segundo   (semilla fija, reproducibles)
"""
import re
import sys
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CSV = sys.argv[1] if len(sys.argv) > 1 else "ReporteRepda.csv"
SALIDA = sys.argv[2] if len(sys.argv) > 2 else "Pozos_Riego_15000.xlsx"
N = 15_000
SEMILLA = 2026

# Clave de 3 letras que aparece dentro del número de título -> entidad federativa
ESTADOS = {
    "AGS": "Aguascalientes", "BCA": "Baja California", "BCS": "Baja California Sur", "CAM": "Campeche",
    "CHS": "Chiapas", "CHI": "Chihuahua", "COA": "Coahuila", "COL": "Colima", "DGO": "Durango",
    "GUA": "Guanajuato", "GRO": "Guerrero", "HGO": "Hidalgo", "JAL": "Jalisco", "MEX": "Estado de México",
    "MCH": "Michoacán", "MOR": "Morelos", "NAY": "Nayarit", "NVL": "Nuevo León", "OAX": "Oaxaca",
    "PUE": "Puebla", "QRO": "Querétaro", "QNR": "Quintana Roo", "SLP": "San Luis Potosí",
    "SIN": "Sinaloa", "SON": "Sonora", "TAB": "Tabasco", "TAM": "Tamaulipas", "TLX": "Tlaxcala",
    "VER": "Veracruz", "YUC": "Yucatán", "ZAC": "Zacatecas",
}

# Municipios agrícolas representativos por estado (catálogo para el campo simulado "municipio")
MUNICIPIOS = {
    "Aguascalientes": ["Pabellón de Arteaga", "Calvillo", "Rincón de Romos", "Jesús María"],
    "Baja California": ["Mexicali", "Ensenada", "Tecate", "San Quintín"],
    "Baja California Sur": ["Comondú", "La Paz", "Mulegé", "Los Cabos"],
    "Campeche": ["Hopelchén", "Champotón", "Calkiní", "Candelaria"],
    "Chiapas": ["Villaflores", "Tapachula", "Villa Corzo", "Frontera Comalapa"],
    "Chihuahua": ["Delicias", "Cuauhtémoc", "Camargo", "Jiménez"],
    "Coahuila": ["Torreón", "Francisco I. Madero", "San Pedro", "Matamoros"],
    "Colima": ["Tecomán", "Armería", "Colima", "Coquimatlán"],
    "Durango": ["Gómez Palacio", "Lerdo", "Tlahualilo", "Santiago Papasquiaro"],
    "Guanajuato": ["Celaya", "Irapuato", "Salamanca", "Valle de Santiago"],
    "Guerrero": ["Iguala", "Tecpan de Galeana", "Coyuca de Benítez", "Cocula"],
    "Hidalgo": ["Ixmiquilpan", "Tula de Allende", "Actopan", "Francisco I. Madero"],
    "Jalisco": ["Autlán de Navarro", "Ameca", "Ojuelos de Jalisco", "Tequila"],
    "Estado de México": ["Jilotepec", "Atlacomulco", "Texcoco", "Tenancingo"],
    "Michoacán": ["Zamora", "Apatzingán", "Buenavista", "La Piedad"],
    "Morelos": ["Cuautla", "Yautepec", "Ayala", "Jojutla"],
    "Nayarit": ["Santiago Ixcuintla", "Tepic", "Ruiz", "Compostela"],
    "Nuevo León": ["Linares", "Montemorelos", "General Terán", "Galeana"],
    "Oaxaca": ["Tlacolula", "Ejutla de Crespo", "Ocotlán", "Santiago Juxtlahuaca"],
    "Puebla": ["Tehuacán", "Tecamachalco", "Atlixco", "Huejotzingo"],
    "Querétaro": ["San Juan del Río", "Pedro Escobedo", "Colón", "Querétaro"],
    "Quintana Roo": ["Othón P. Blanco", "Felipe Carrillo Puerto", "José María Morelos", "Bacalar"],
    "San Luis Potosí": ["Villa de Arista", "Matehuala", "Rioverde", "Ciudad Fernández"],
    "Sinaloa": ["Culiacán", "Guasave", "Ahome", "Navolato"],
    "Sonora": ["Hermosillo", "Cajeme", "Caborca", "Navojoa"],
    "Tabasco": ["Cárdenas", "Huimanguillo", "Balancán", "Centla"],
    "Tamaulipas": ["Río Bravo", "Reynosa", "El Mante", "Valle Hermoso"],
    "Tlaxcala": ["Calpulalpan", "Huamantla", "Tlaxco", "Cuapiaxtla"],
    "Veracruz": ["Tuxpan", "Cosamaloapan", "Pánuco", "Martínez de la Torre"],
    "Yucatán": ["Tizimín", "Tekax", "Peto", "Valladolid"],
    "Zacatecas": ["Fresnillo", "Calera", "Sombrerete", "Villa de Cos"],
}

ESTATUS = ["Activo", "Mantenimiento", "Suspendido", "Clausurado"]
P_ESTATUS = [0.80, 0.08, 0.08, 0.04]


def num(serie):
    """'579 300.00' -> 579300.0"""
    return pd.to_numeric(serie.str.replace(r"[\s,]", "", regex=True), errors="coerce")


def main():
    df = pd.read_csv(CSV, encoding="utf-8-sig", dtype=str)
    df.columns = ["titular", "titulo", "uso", "autoridad", "fecha", "vol_ext", "n_sup", "v_sup",
                  "n_sub", "v_sub", "n_desc", "v_desc", "n_zf", "sup"]
    print(f"Registros en CSV: {len(df):,}")

    # 1) Filtro: uso agrícola con aprovechamiento de aguas subterráneas (pozos)
    df = df[df["uso"].str.strip().str.upper().eq("AGRICOLA")]
    df = df[(num(df["v_sub"]) > 0) & (num(df["vol_ext"]) > 0)]

    # 2) id_titulo_concesion: folio numérico de 6 dígitos del título (04MCH162608/18AMDL26 -> 162608)
    ext = df["titulo"].str.extract(r"^[0-9A-Z_]{1,2}([A-Z]{3})(\d{6})/")
    df = df.assign(cve=ext[0], folio=ext[1])
    df = df[df["cve"].isin(ESTADOS) & df["folio"].notna()]

    # 3) Fecha válida -> entero AAAAMMDD
    f = pd.to_datetime(df["fecha"], format="%d/%m/%Y", errors="coerce")
    df = df[f.notna()]
    f = f[f.notna()]

    df = df.assign(
        id_titulo_concesion=df["folio"].astype(int),
        volumen_extraido_m3=num(df["vol_ext"]).round().astype(int),
        fecha_concesion=f.dt.strftime("%Y%m%d").astype(int),
        nombre_titular=df["titular"].fillna("").str.strip().str.replace(r"\s{2,}", " ", regex=True).str.slice(0, 80),
        estado=df["cve"].map(ESTADOS),
        uso_principal="Agrícola",
    )
    df = df[df["nombre_titular"] != ""]
    # El folio debe ser único (campo clave)
    df = df.drop_duplicates("id_titulo_concesion")
    print(f"Candidatos válidos (únicos): {len(df):,}")

    # 4) Muestra reproducible de 15,000
    rng = np.random.default_rng(SEMILLA)
    df = df.sample(N, random_state=SEMILLA).reset_index(drop=True)

    # 5) Campos simulados (no existen en el REPDA)
    df["municipio"] = [rng.choice(MUNICIPIOS[e]) for e in df["estado"]]
    df["profundidad_pozo_m"] = np.clip(rng.lognormal(np.log(110), 0.55, N), 15, 450).round(1)
    df["estatus_pozo"] = rng.choice(ESTATUS, N, p=P_ESTATUS)
    # Gasto medio estimado si el pozo operara ~12 h/día: Q = V / (365 d * 12 h * 3600 s) * 1000 L/m3
    q = df["volumen_extraido_m3"] * 1000 / (365 * 12 * 3600) * rng.normal(1.0, 0.12, N)
    df["caudal_litros_segundo"] = np.clip(q, 0.1, None).round(1)

    cols = ["id_titulo_concesion", "volumen_extraido_m3", "fecha_concesion", "nombre_titular", "estado",
            "municipio", "profundidad_pozo_m", "estatus_pozo", "caudal_litros_segundo", "uso_principal"]
    out = df[cols]
    escribir_excel(out, SALIDA)
    out.to_csv(SALIDA.replace(".xlsx", ".csv"), index=False, encoding="utf-8-sig")
    print(f"Listo: {SALIDA}")
    print(out.head(5).to_string())


def escribir_excel(d, ruta):
    wb = Workbook()
    azul = PatternFill("solid", fgColor="1F4E78")
    verde = PatternFill("solid", fgColor="E2EFDA")
    ambar = PatternFill("solid", fgColor="FFF2CC")
    blanco = Font(bold=True, color="FFFFFF")

    # --- Hoja Datos
    ws = wb.active
    ws.title = "Datos"
    ws.append(list(d.columns))
    for fila in d.itertuples(index=False):
        ws.append([x.item() if hasattr(x, "item") else x for x in fila])
    for c in ws[1]:
        c.font, c.fill = blanco, azul
        c.alignment = Alignment(horizontal="center")
    for i, w in enumerate([22, 22, 18, 45, 20, 24, 20, 16, 22, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # --- Hoja Registro_Ejecucion (formato del documento, para llenar a mano)
    r = wb.create_sheet("Registro_Ejecucion")
    r.append(["Ejecución #", "Campo Clave Seleccionado (1-3)", "Nombre del Campo",
              "Método de Ordenamiento Asignado", "Tiempo de Ejecución (ns)"])
    for i in (1, 2, 3):
        r.append([i, None, None, None, None])
    for c in r[1]:
        c.font, c.fill = blanco, azul
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, w in enumerate([12, 30, 26, 34, 26], 1):
        r.column_dimensions[get_column_letter(i)].width = w

    # --- Hoja Mapeo_Campos
    m = wb.create_sheet("Mapeo_Campos")
    m.append(["#", "Campo requerido", "Tipo", "Rol", "Columna origen en ReporteRepda.csv", "Origen", "Transformación"])
    filas = [
        (1, "id_titulo_concesion", "Entero", "CLAVE 1 (Radix / Counting)", "Título", "REAL", "Folio de 6 dígitos del título: 04MCH162608/18AMDL26 -> 162608"),
        (2, "volumen_extraido_m3", "Entero", "CLAVE 2 (Bucket)", "Volumen de extracción de aguas nacionales (m3/año)", "REAL", "Se quitan espacios ('579 300.00') y se redondea a entero"),
        (3, "fecha_concesion", "Entero AAAAMMDD", "CLAVE 3 (Quick / Merge / Heap)", "Fecha de registro", "REAL", "dd/mm/aaaa -> aaaammdd"),
        (4, "nombre_titular", "Texto", "Informativo", "Titular", "REAL", "Se limpian espacios; máx. 80 caracteres"),
        (5, "estado", "Texto", "Informativo", "Título (clave de 3 letras: MCH, GUA, CHI...)", "REAL", "Clave -> entidad federativa"),
        (6, "municipio", "Texto", "Informativo", "No existe en el CSV", "SIMULADO", "Municipio agrícola del catálogo del estado, al azar (semilla 2026)"),
        (7, "profundidad_pozo_m", "Flotante", "Informativo", "No existe en el CSV", "SIMULADO", "Lognormal (mediana 110 m, rango 15–450 m), 1 decimal"),
        (8, "estatus_pozo", "Texto", "Informativo", "No existe en el CSV", "SIMULADO", "Activo 80% / Mantenimiento 8% / Suspendido 8% / Clausurado 4%"),
        (9, "caudal_litros_segundo", "Flotante", "Informativo", "No existe en el CSV", "SIMULADO", "Derivado del volumen anual suponiendo 12 h/día de bombeo, ±12% de ruido"),
        (10, "uso_principal", "Texto", "Informativo", "Uso", "REAL", "Se filtra Uso = AGRICOLA"),
    ]
    for f in filas:
        m.append(f)
    for c in m[1]:
        c.font, c.fill = blanco, azul
    for row in m.iter_rows(min_row=2):
        for c in row:
            c.fill = verde if row[5].value == "REAL" else ambar
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([5, 24, 18, 28, 46, 12, 62], 1):
        m.column_dimensions[get_column_letter(i)].width = w
    wb.save(ruta)


if __name__ == "__main__":
    main()
