"""
Consolidación de Valorización SIS · Red de Salud Satipo
---------------------------------------------------------
Lee:
  1) Los Excel de VALORIZACIÓN (los que exportas de tu carpeta de
     valorización — se pueden ir acumulando, este script deduplica por FUA
     y se queda con la fila más reciente de cada uno).
  2) Los Excel RAtenDet de esa misma carpeta de "BASE DE DATOS 2" —
     de ahí se saca el DNI y nombre del PROFESIONAL DE SALUD que atendió
     cada FUA (columnas "DNI Resp. Aten." / "Profesional" / "Tipo
     Profesional"). El campo "Digitador" también se guarda por si acaso,
     pero no es el foco de este dashboard.
  3) Cruza ambas fuentes por N° de FUA y escribe un data/val_AAAA_MM.json
     por cada mes de valorización (Periodo+Mes), más un manifest.json.

Además marca cada atención con "diagnostico_percapita": True si su
diagnóstico está en el Anexo de diagnósticos con cobertura por mecanismo
de pago per cápita (esas atenciones no generan valorización adicional
por fuera de la cápita).

Uso:
    python consolidar_valorizacion.py

Requisitos:
    pip install pandas python-calamine
"""

import json
import re
import unicodedata
from pathlib import Path
from datetime import datetime

import pandas as pd

# ============================================================================
# 1) CONFIGURA TUS CARPETAS AQUÍ
# ============================================================================

# Carpeta donde caen los Excel de valorización (uno o varios, se acumulan).
VALORIZ_DIR = Path(r"D:\ANALISIS DE DATOS\VALORIZACION BASE DE DATOS")

# Carpeta con los RAtenDet_*.xlsx (misma que usas para Producción del
# Digitador / Personal de Salud). También puedes apuntar aquí directamente
# a la carpeta data\ del proyecto "producción-digitador" y dejar sus
# dig_AAAA_MM.json ya consolidados — el script los detecta solo.
RATENDET_DIR = Path(r"C:\Users\USUARIO\Desktop\BASE DE DATOS\Escritorio\USPP\BASE DE DATOS 2")

OUT_DIR = Path(__file__).parent / "data"

# Columnas que NUNCA se guardan (datos de identificación del paciente).
# El nombre y DNI del profesional/digitador SÍ se guardan (igual que en el
# proyecto de Producción), porque ese es justo el dato que se quiere cruzar.
COLUMNAS_PACIENTE_EXCLUIR = True

# ============================================================================


# Anexo: diagnósticos CIE10 que, al usarse, quedan cubiertos por el
# mecanismo de pago per cápita (no generan valorización adicional por
# fuera de la cápita — la atención "se pierde" para efectos de pago FFS).
# 354 códigos, tal como el Anexo compartido.
DIAGNOSTICOS_PERCAPITA = {
    "A049", "A059", "A069", "A071", "A084", "A090", "A09X", "A150", "A162", "A169",
    "A539", "A64X", "A970", "B009", "B019", "B029", "B084", "B24X", "B349", "B351",
    "B353", "B354", "B359", "B369", "B373", "B49X", "B54X", "B551", "B779", "B80X",
    "B829", "B86X", "B889", "B968", "D508", "D509", "D539", "D649", "E440", "E45X",
    "E46X", "E617", "E631", "E639", "E660", "E669", "E780", "E781", "E782", "E784",
    "E785", "E788", "E789", "I872", "G409", "G430", "G431", "G439", "G442", "H000",
    "H100", "H101", "H103", "H109", "H110", "H259", "H527", "H571", "H609", "H612",
    "H650", "H651", "H660", "H669", "H811", "H813", "H920", "J00X", "J019", "J020",
    "J028", "J029", "J030", "J038", "J039", "J040", "J041", "J042", "J060", "J068",
    "J069", "J111", "J159", "J189", "J209", "J219", "J303", "J304", "J40X", "J448",
    "J450", "J459", "K020", "K021", "K029", "K036", "K040", "K041", "K046", "K047",
    "K050", "K083", "K088", "K120", "K122", "K219", "K291", "K293", "K295", "K296",
    "K297", "K30X", "K529", "K590", "K591", "K649", "K759", "K760", "K769", "K808",
    "L010", "L020", "L024", "L028", "L029", "L030", "L031", "L038", "L039", "L080",
    "L089", "L208", "L209", "L22X", "L239", "L259", "L299", "L303", "L309", "L500",
    "L509", "L600", "M139", "M150", "M159", "M170", "M179", "M199", "M245", "M255",
    "M419", "M511", "M542", "M544", "M545", "M546", "M548", "M549", "M624", "M629",
    "M779", "M791", "M792", "M796", "M819", "N200", "N209", "N23X", "N300", "N309",
    "N390", "N40X", "N61X", "N644", "N72X", "N739", "N760", "N761", "N762", "N771",
    "N832", "N939", "N946", "N952", "S001", "S009", "S010", "S011", "S012", "S015",
    "S018", "S019", "S300", "S311", "S400", "S411", "S519", "S600", "S602", "S610",
    "S611", "S618", "S619", "S700", "S701", "S711", "S800", "S801", "S810", "S817",
    "S818", "S819", "S900", "S903", "S910", "S911", "S912", "S913", "S917", "S934",
    "T140", "T141", "T749", "T784", "T793", "O149", "O210", "O233", "O234", "O239",
    "O260", "O261", "O800", "O809", "O990", "R040", "R05X", "R070", "R074", "R101",
    "R102", "R103", "R104", "R11X", "R12X", "R14X", "R33X", "R42X", "R500", "R501",
    "R509", "R51X", "R520", "R522", "R529", "R53X", "R633", "R739", "W540", "W544",
    "W549", "W570", "Z000", "Z001", "Z003", "Z006", "Z008", "Z010", "Z012", "Z014",
    "Z016", "Z017", "Z018", "Z019", "Z030", "Z049", "Z098", "Z111", "Z113", "Z114",
    "Z116", "Z121", "Z123", "Z125", "Z128", "Z129", "Z132", "Z133", "Z138", "Z201",
    "Z238", "Z240", "Z243", "Z246", "Z251", "Z258", "Z268", "Z269", "Z271", "Z274",
    "Z278", "Z279", "Z298", "Z300", "Z304", "Z308", "Z309", "Z320", "Z321", "Z340",
    "Z348", "Z349", "Z352", "Z357", "Z358", "Z359", "Z368", "Z369", "Z370", "Z392",
    "Z480", "Z501", "Z507", "Z519", "Z588", "Z591", "Z594", "Z630", "Z634", "Z637",
    "Z653", "Z713", "Z718", "Z719", "Z721", "Z724", "Z728", "Z739", "Z742", "Z743",
    "Z762", "Z768", "Z918", "Z928",
}


def normaliza(s: str) -> str:
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)


def encontrar_col(columnas, candidatos):
    mapa = {normaliza(c): c for c in columnas}
    for cand in candidatos:
        if normaliza(cand) in mapa:
            return mapa[normaliza(cand)]
    return None


def extraer_timestamp_archivo(nombre: str):
    """'RAtenDet_07-09-2026_11-52-36CANAAN.xlsx' -> datetime(2026,9,7,11,52,36)"""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})_(\d{2})-(\d{2})-(\d{2})", nombre)
    if not m:
        return None
    dd, mm, yyyy, hh, mi, ss = map(int, m.groups())
    try:
        return datetime(yyyy, mm, dd, hh, mi, ss)
    except ValueError:
        return None


def orden_archivos(paths):
    """Ordena cronológicamente usando la fecha/hora del nombre si existe,
    y si no, la fecha de modificación del archivo. Así, al deduplicar por
    FUA, siempre gana la versión más reciente (los RAtenDet/valorización
    del SIS reexportan todo el histórico cada vez)."""

    def clave(p):
        ts = extraer_timestamp_archivo(p.name)
        return ts or datetime.fromtimestamp(p.stat().st_mtime)

    return sorted(paths, key=clave)


def fecha_a_iso(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, str):
        valor = valor.strip()
        if not valor:
            return None
        for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(valor, fmt).date().isoformat()
            except ValueError:
                continue
        return valor
    try:
        return pd.Timestamp(valor).date().isoformat()
    except Exception:
        return str(valor)


def hora_de(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, str):
        valor = valor.strip()
        m = re.search(r"(\d{1,2}):(\d{2})", valor)
        return m.group(0) if m else (valor or None)
    try:
        return pd.Timestamp(valor).strftime("%H:%M")
    except Exception:
        return None


# ============================================================================
# 2) LECTURA DE VALORIZACIÓN
# ============================================================================

CAND_VALORIZ = {
    "fua": ["FUA"],
    "periodo": ["Periodo"],
    "mes": ["Mes"],
    "eess": ["CodigoEESS"],
    "categoria": ["Categoria"],
    "servicio_id": ["Servicio"],
    "servicio": ["ServicioDesc"],
    "profesion": ["PROFESION"],
    "especialidad": ["ESPECIALISTA"],
    "observado": ["Observado_PSA"],
    "capita": ["Capita"],
    "fecha_atencion": ["ate_fecatencion"],
    "edad": ["ate_edad"],
    "sexo": ["ate_idsexo"],
    "diagnostico": ["ate_coddiag1"],
    "valor_bruto_sis": ["ValorBruto_SIS"],
    "valor_bruto_serv_sis": ["ValorBrutoSer_SIS"],
    "valor_bruto_med_sis": ["ValorBrutoMed_SIS"],
    "valor_bruto_ins_sis": ["ValorBrutoIns_SIS"],
    "valor_bruto_apo_sis": ["ValorBrutoApo_SIS"],
    "valor_neto_sis": ["ValorNeto_SIS"],
}


def _leer_hoja_con_fua(path):
    """Los Excel de valorización a veces traen una hoja de tabla dinámica
    (resumen) antes de la hoja de detalle. Esta función prueba todas las
    hojas del archivo y devuelve la primera que tenga columna FUA."""
    try:
        hojas = pd.ExcelFile(path, engine="calamine").sheet_names
    except Exception:
        hojas = [0]
    for hoja in hojas:
        try:
            df = pd.read_excel(path, sheet_name=hoja, engine="calamine")
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        if encontrar_col(df.columns, CAND_VALORIZ["fua"]):
            return df
    return None


def leer_valorizacion(paths):
    """Lee y deduplica por FUA (se queda con la fila del archivo más
    reciente). Devuelve dict FUA -> registro."""
    registros = {}
    for path in orden_archivos(paths):
        df = _leer_hoja_con_fua(path)
        if df is None:
            print(f"  ! {path.name}: no se encontró columna FUA en ninguna hoja, se omite")
            continue
        cols = {clave: encontrar_col(df.columns, cands) for clave, cands in CAND_VALORIZ.items()}

        for _, row in df.iterrows():
            fua = row.get(cols["fua"])
            if pd.isna(fua):
                continue
            fua = str(fua).strip()

            def val(clave):
                c = cols.get(clave)
                if not c:
                    return None
                v = row.get(c)
                return None if (v is None or (isinstance(v, float) and pd.isna(v))) else v

            registros[fua] = {
                "fua": fua,
                "periodo": str(val("periodo")) if val("periodo") is not None else None,
                "mes": str(val("mes")).zfill(2) if val("mes") is not None else None,
                "eess": val("eess"),
                "categoria": val("categoria"),
                "servicio_id": val("servicio_id"),
                "servicio": val("servicio"),
                "profesion": val("profesion"),
                "especialidad": val("especialidad"),
                "observado": val("observado"),
                "capita": val("capita"),
                "fecha_atencion": fecha_a_iso(val("fecha_atencion")),
                "hora_atencion": hora_de(val("fecha_atencion")),
                "edad": val("edad"),
                "sexo": val("sexo"),
                "diagnostico": val("diagnostico"),
                "diagnostico_percapita": (str(val("diagnostico")).strip().upper() in DIAGNOSTICOS_PERCAPITA)
                if val("diagnostico") is not None else False,
                "valor_bruto_sis": float(val("valor_bruto_sis") or 0),
                "valor_bruto_serv_sis": float(val("valor_bruto_serv_sis") or 0),
                "valor_bruto_med_sis": float(val("valor_bruto_med_sis") or 0),
                "valor_bruto_ins_sis": float(val("valor_bruto_ins_sis") or 0),
                "valor_bruto_apo_sis": float(val("valor_bruto_apo_sis") or 0),
                "valor_neto_sis": float(val("valor_neto_sis") or 0),
            }
            r = registros[fua]
            if r["valor_bruto_sis"] > 0:
                r["motivo_no_valorizado"] = None
            elif r["diagnostico_percapita"]:
                r["motivo_no_valorizado"] = "per_capita"
            elif not r["especialidad"]:
                # Por el convenio: sin médico especialista (RNE) registrado en el
                # FUA, la prestación tampoco se valoriza aunque el diagnóstico
                # no esté en el Anexo 19.
                r["motivo_no_valorizado"] = "sin_especialista"
            else:
                r["motivo_no_valorizado"] = "sin_explicacion"
        print(f"  ✓ {path.name}: {len(df):,} filas leídas")
    return registros


# ============================================================================
# 3) LECTURA DE RAtenDet (producción / digitador) — Excel crudo o JSON ya
#    consolidado por el proyecto de Producción del Digitador.
# ============================================================================

CAND_RATENDET = {
    "fua": ["FUA"],
    "dni_resp_atencion": ["DNI Resp. Aten.", "DNI Resp Aten"],
    "profesional_nombre": ["Profesional"],
    "profesional_tipo": ["Tipo Profesional"],
    "digitador_dni": ["Digitador"],
    "fecha_registro": ["Fecha Registro"],
    "hora_registro": ["Hora Registro"],
    "ppdd": ["PPDD", "Punto Digitacion", "Punto de Digitacion"],
    "eess_registro": ["EESS"],
    "estado_fua": ["Estado FUA"],
    "importe_referencial": ["Importe Referencial"],
}


def leer_ratendet_excel(path):
    df = pd.read_excel(path, engine="calamine")
    df.columns = [str(c).strip() for c in df.columns]
    cols = {clave: encontrar_col(df.columns, cands) for clave, cands in CAND_RATENDET.items()}
    if not cols["fua"] or not (cols["dni_resp_atencion"] or cols["digitador_dni"]):
        print(f"  ! {path.name}: no tiene columnas FUA / DNI Resp. Aten., se omite")
        return {}

    registros = {}
    for _, row in df.iterrows():
        fua = row.get(cols["fua"])
        if pd.isna(fua):
            continue
        fua = str(fua).strip()

        def val(clave):
            c = cols.get(clave)
            if not c:
                return None
            v = row.get(c)
            return None if (v is None or (isinstance(v, float) and pd.isna(v))) else v

        registros[fua] = {
            "dni_resp_atencion": str(val("dni_resp_atencion")).strip() if val("dni_resp_atencion") is not None else None,
            "profesional_nombre": val("profesional_nombre"),
            "profesional_tipo": val("profesional_tipo"),
            "digitador_dni": str(val("digitador_dni")).strip() if val("digitador_dni") is not None else None,
            "fecha_registro": fecha_a_iso(val("fecha_registro")),
            "hora_registro": hora_de(val("hora_registro")),
            "ppdd": val("ppdd"),
            "eess_registro": val("eess_registro"),
            "estado_fua": val("estado_fua"),
            "importe_referencial": float(val("importe_referencial") or 0),
        }
    return registros


def leer_ratendet_json(path):
    """Lee un dig_AAAA_MM.json ya consolidado por el proyecto de
    Producción del Digitador y lo deja en el mismo formato que
    leer_ratendet_excel(), para poder mezclar ambas fuentes. Ese proyecto
    no guarda el profesional de salud (solo el digitador), así que esos
    campos quedan en None cuando la fuente es un .json."""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    registros = {}
    for r in data:
        fua = str(r.get("fua", "")).strip()
        if not fua:
            continue
        registros[fua] = {
            "dni_resp_atencion": r.get("dni_resp_atencion"),
            "profesional_nombre": r.get("profesional_nombre"),
            "profesional_tipo": r.get("profesional_tipo") or r.get("tipo_profesional"),
            "digitador_dni": r.get("digitador_dni") or r.get("dni"),
            "fecha_registro": r.get("fecha_registro"),
            "hora_registro": r.get("hora_registro"),
            "ppdd": r.get("punto_digitacion") or r.get("ppdd"),
            "eess_registro": r.get("establecimiento") or r.get("eess"),
            "estado_fua": r.get("estado_fua"),
            "importe_referencial": r.get("importe_referencial"),
        }
    return registros


def leer_ratendet(paths):
    registros = {}
    for path in orden_archivos(paths):
        if path.suffix.lower() == ".json":
            nuevos = leer_ratendet_json(path)
            print(f"  ✓ {path.name}: {len(nuevos):,} FUA (json ya consolidado)")
        else:
            nuevos = leer_ratendet_excel(path)
            print(f"  ✓ {path.name}: {len(nuevos):,} filas leídas")
        registros.update(nuevos)  # el más reciente (por orden cronológico) gana
    return registros


# ============================================================================
# 4) CRUCE + EXPORTACIÓN
# ============================================================================


def dias_entre(fecha_ini, fecha_fin):
    if not fecha_ini or not fecha_fin:
        return None
    try:
        d1 = datetime.fromisoformat(fecha_ini)
        d2 = datetime.fromisoformat(fecha_fin)
        return (d2 - d1).days
    except ValueError:
        return None


def cruzar_y_guardar(valorizacion: dict, ratendet: dict):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    buckets = {}  # (periodo, mes) -> [registros]
    encontrados = 0

    for fua, v in valorizacion.items():
        prod = ratendet.get(fua)
        registro = dict(v)
        if prod and (prod.get("dni_resp_atencion") or prod.get("digitador_dni")):
            encontrados += 1
            registro.update(
                {
                    "dni_resp_atencion": prod["dni_resp_atencion"],
                    "profesional_nombre": prod["profesional_nombre"],
                    "profesional_tipo": prod["profesional_tipo"],
                    "digitador_dni": prod["digitador_dni"],
                    "fecha_registro": prod["fecha_registro"],
                    "hora_registro": prod["hora_registro"],
                    "ppdd": prod["ppdd"],
                    "eess_registro": prod["eess_registro"],
                    "estado_fua": prod["estado_fua"],
                    "lag_dias": dias_entre(v["fecha_atencion"], prod["fecha_registro"]),
                    "importe_referencial": prod.get("importe_referencial") or 0,
                }
            )
        else:
            registro.update(
                {
                    "dni_resp_atencion": None,
                    "profesional_nombre": None,
                    "profesional_tipo": None,
                    "digitador_dni": None,
                    "fecha_registro": None,
                    "hora_registro": None,
                    "ppdd": None,
                    "eess_registro": None,
                    "estado_fua": None,
                    "lag_dias": None,
                    "importe_referencial": 0,
                }
            )

        periodo = registro.get("periodo") or "0000"
        mes = registro.get("mes") or "00"
        buckets.setdefault((periodo, mes), []).append(registro)

    disponibles = []
    for (periodo, mes), regs in sorted(buckets.items()):
        nombre = f"val_{periodo}_{mes}.json"
        with open(OUT_DIR / nombre, "w", encoding="utf-8") as f:
            json.dump(regs, f, ensure_ascii=False, separators=(",", ":"))
        total_sis = sum(r["valor_bruto_sis"] for r in regs)
        disponibles.append({"anio": int(periodo), "mes": int(mes), "file": nombre})
        print(f"  → {nombre}: {len(regs):,} atenciones, S/ {total_sis:,.0f}")

    manifest = {
        "disponibles": disponibles,
        "ultimo": disponibles[-1] if disponibles else None,
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "origen_valorizacion": str(VALORIZ_DIR),
        "origen_produccion": str(RATENDET_DIR),
    }
    with open(OUT_DIR / "manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    todas = [r for regs in buckets.values() for r in regs]
    total_regs = len(todas)
    total_valor = sum(r["valor_bruto_sis"] for r in todas)
    valorizadas = [r for r in todas if r["valor_bruto_sis"] > 0]
    percap = [r for r in todas if r["motivo_no_valorizado"] == "per_capita"]
    sin_esp = [r for r in todas if r["motivo_no_valorizado"] == "sin_especialista"]
    sin_exp = [r for r in todas if r["motivo_no_valorizado"] == "sin_explicacion"]
    potencial_sin_exp = sum(r["importe_referencial"] for r in sin_exp)
    potencial_sin_esp = sum(r["importe_referencial"] for r in sin_esp)

    print(f"\nTotal FUA valorización: {len(valorizacion):,}")
    print(f"Con profesional de salud identificado (cruce por FUA): {encontrados:,} ({100*encontrados/max(len(valorizacion),1):.1f}%)")
    print(f"Meses generados: {len(buckets)}")
    print(f"\nValor Bruto SIS total: S/ {total_valor:,.0f}")
    print(f"  1) Valorizadas (>0):                 {len(valorizadas):6,d} ({100*len(valorizadas)/max(total_regs,1):.1f}%)")
    print(f"  2) Cero, per cápita (Anexo 19):       {len(percap):6,d} ({100*len(percap)/max(total_regs,1):.1f}%)")
    print(f"  3) Cero, sin especialista (RNE):      {len(sin_esp):6,d} ({100*len(sin_esp)/max(total_regs,1):.1f}%)")
    print(f"  4) Cero, SIN explicación (revisar):   {len(sin_exp):6,d} ({100*len(sin_exp)/max(total_regs,1):.1f}%)")
    print(f"\nSimulación con Importe Referencial (RAtenDet):")
    print(f"  Si se valorizara el grupo 4 (sin explicación): +S/ {potencial_sin_exp:,.0f}")
    print(f"  Si además se registrara especialista en el grupo 3: +S/ {potencial_sin_esp:,.0f}")
    print(f"  Potencial total adicional: S/ {potencial_sin_exp + potencial_sin_esp:,.0f}"
          f"  (nuevo total simulado: S/ {total_valor + potencial_sin_exp + potencial_sin_esp:,.0f})")


# ============================================================================
# 5) MAIN
# ============================================================================

if __name__ == "__main__":
    print("1) Leyendo Excel de VALORIZACIÓN...")
    archivos_valoriz = sorted(VALORIZ_DIR.glob("*.xlsx")) if VALORIZ_DIR.exists() else []
    if not archivos_valoriz:
        print(f"   No se encontraron .xlsx en {VALORIZ_DIR}")
    valorizacion = leer_valorizacion(archivos_valoriz)

    print("\n2) Leyendo RAtenDet / producción (Excel o JSON ya consolidado)...")
    archivos_ratendet = []
    if RATENDET_DIR.exists():
        archivos_ratendet = sorted(RATENDET_DIR.glob("*.xlsx")) + sorted(RATENDET_DIR.glob("*.json"))
    if not archivos_ratendet:
        print(f"   No se encontraron archivos en {RATENDET_DIR}")
    ratendet = leer_ratendet(archivos_ratendet)

    print("\n3) Cruzando por FUA y guardando...")
    cruzar_y_guardar(valorizacion, ratendet)
