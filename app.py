import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import hashlib
from copy import copy
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import plotly.graph_objects as go
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

# ==============================================================================
# CONFIGURACIÓN
# ==============================================================================
APP_VERSION = "5.3 ARPON · HOTEL QUARTZ · CLIENTES + PROVEEDORES"
# Tolerancia contable en moneda. Se usa para validaciones y conciliaciones.
# Dos centavos evitan falsos errores binarios de float sin aceptar diferencias materiales.
UMBRAL_TOLERANCIA = 0.02
UMBRAL_FOLIO = 0.01

# Prefijos documentales que sí tratamos como folios.
# Se conservan; NO se eliminan durante la normalización.
PREFIJOS_FOLIO = (
    # Prefijos documentales observados en auxiliares ARPON de Hotel Quartz.
    "NCTA", "NC", "H", "B", "R", "X", "E", "S",
)

REFERENCIAS_VACIAS = {
    "", "N/A", "NA", "N.A.", "SIN REF", "SIN REFERENCIA", "S/R",
    "NO APLICA", "NO APLICA.", "NINGUNA", "-", "—", "–", "0",
}

MESES_ES = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "aug": 8, "sep": 9, "sept": 9, "oct": 10,
    "nov": 11, "dic": 12, "dec": 12, "jan": 1, "apr": 4,
}

# ==============================================================================
# 1. UTILIDADES
# ==============================================================================

def quitar_acentos(texto):
    texto = "" if texto is None else str(texto)
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )


def texto_norm(texto):
    if pd.isna(texto):
        return ""
    s = quitar_acentos(str(texto)).upper().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def concepto_norm(texto):
    """Normalización conservadora para comparar conceptos entre cuentas."""
    s = texto_norm(texto)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_spanish_date(valor):
    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, (pd.Timestamp, np.datetime64)):
        return pd.Timestamp(valor)

    # Excel puede entregar la fecha como serial numérico, sobre todo en CSV
    # exportados o libros con formatos poco consistentes.
    if isinstance(valor, (int, float, np.integer, np.floating)):
        n = float(valor)
        if 20000 <= n <= 80000:
            try:
                return pd.Timestamp("1899-12-30") + pd.to_timedelta(n, unit="D")
            except Exception:
                return pd.NaT

    s = str(valor).strip()
    if not s:
        return pd.NaT

    # Serial de Excel guardado como texto.
    if re.fullmatch(r"\d{5}(?:\.\d+)?", s):
        n = float(s)
        if 20000 <= n <= 80000:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(n, unit="D")

    # dd/Mmm/aaaa, aceptando abreviaturas ES/EN.
    m = re.match(
        r"^(\d{1,2})[/\-]([A-Za-zÁÉÍÓÚáéíóúÑñ]{3,4})[/\-](\d{4})$",
        s,
    )
    if m:
        day, mon_abbr, year = m.groups()
        mon = quitar_acentos(mon_abbr).lower()
        if mon in MESES_ES:
            try:
                return pd.Timestamp(
                    year=int(year), month=MESES_ES[mon], day=int(day)
                )
            except ValueError:
                return pd.NaT

    # ISO yyyy-mm-dd / yyyy-mm-dd hh:mm:ss.
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s):
        return pd.to_datetime(s, errors="coerce")

    # Evita interpretar números aislados (por ejemplo "2026") como fechas.
    if not re.search(r"[/\-]", s) and not re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", s):
        return pd.NaT

    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def parse_amount(valor):
    """
    Convierte montos contables sin convertir silenciosamente texto inválido en cero.
    Soporta:
      1,234.56
      $1,234.56
      (1,234.56)
      -1,234.56
      vacío -> 0
    """
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float, np.integer, np.floating)):
        if pd.isna(valor):
            return 0.0
        return float(valor)

    s = str(valor).strip()
    if s == "":
        return 0.0

    negativo_parentesis = s.startswith("(") and s.endswith(")")
    if negativo_parentesis:
        s = s[1:-1].strip()

    s = s.replace("$", "").replace(",", "").replace(" ", "")
    s = s.replace("MXN", "").replace("USD", "")

    # Guiones aislados suelen representar vacío.
    if s in {"-", "—", "–"}:
        return 0.0

    try:
        n = float(s)
        return -n if negativo_parentesis else n
    except ValueError:
        return np.nan


def columna_a_monto(serie):
    return serie.apply(parse_amount)


def es_vacio(valor):
    return pd.isna(valor) or str(valor).strip() == ""



def detectar_tipo_cuenta(nombre):
    """Clasifica únicamente las familias soportadas por este auditor."""
    s = texto_norm(nombre)
    if "PROVEED" in s:
        return "PROVEEDORES"
    if "CLIENT" in s:
        return "CLIENTES"
    return "NO_SOPORTADA"


def factor_pendiente_tipo(tipo_cuenta):
    """
    Convierte el signo ARPON a una magnitud de pendiente homogénea:
      CLIENTES:    +saldo = por cobrar        -> factor +1
      PROVEEDORES: -saldo = por pagar         -> factor -1
    """
    if tipo_cuenta == "CLIENTES":
        return 1.0
    if tipo_cuenta == "PROVEEDORES":
        return -1.0
    return np.nan


def reconstruir_filas_fragmentadas_arpon(raw, file_name):
    """
    Repara de forma conservadora el patrón observado en exportaciones ARPON donde
    una partida se parte en dos filas:

      fila N:   Póliza | Fecha | Docto. | <vacío> | <vacío> | <vacío> | <vacío>
      fila N+1: <vacío>| Concepto       | Cargo   | Abono   | Saldo   | <vacío> | <vacío>

    La reparación solo mueve Concepto/Cargo/Abono/Saldo a la fila principal.
    La fila de continuación NO se elimina, para conservar la numeración original.
    La validez final queda certificada por los amarres de totales y la secuencia
    completa de saldo; si no cuadran, el archivo se rechaza posteriormente.
    """
    x = raw.copy()
    reparaciones = []
    etiquetas_no_concepto = {"TOTALES", "NETO PERIODO", "POLIZA", "FECHA"}

    for i in range(len(x) - 1):
        r = x.iloc[i]
        n = x.iloc[i + 1]

        principal = (
            not es_vacio(r.iloc[0])
            and pd.notna(parse_spanish_date(r.iloc[1]))
            and not es_vacio(r.iloc[2])
            and all(es_vacio(r.iloc[c]) for c in (3, 4, 5, 6))
        )
        if not principal:
            continue

        concepto_sig = texto_norm(n.iloc[1])
        continuacion = (
            es_vacio(n.iloc[0])
            and not es_vacio(n.iloc[1])
            and pd.isna(parse_spanish_date(n.iloc[1]))
            and concepto_sig not in etiquetas_no_concepto
            and not es_vacio(n.iloc[2])
            and not es_vacio(n.iloc[3])
            and not es_vacio(n.iloc[4])
            and es_vacio(n.iloc[5])
            and es_vacio(n.iloc[6])
        )
        if not continuacion:
            continue

        cargo = parse_amount(n.iloc[2])
        abono = parse_amount(n.iloc[3])
        saldo = parse_amount(n.iloc[4])
        if any(pd.isna(v) for v in (cargo, abono, saldo)):
            continue

        x.iat[i, 3] = n.iloc[1]
        x.iat[i, 4] = float(cargo)
        x.iat[i, 5] = float(abono)
        x.iat[i, 6] = float(saldo)

        reparaciones.append(
            {
                "archivo": file_name,
                "fila_principal": int(i + 1),
                "fila_continuacion": int(i + 2),
                "poliza": str(r.iloc[0]).strip(),
                "fecha": parse_spanish_date(r.iloc[1]),
                "referencia": str(r.iloc[2]).strip(),
                "concepto_reconstruido": str(n.iloc[1]).strip(),
                "cargo_reconstruido": float(cargo),
                "abono_reconstruido": float(abono),
                "saldo_reconstruido": float(saldo),
                "reparacion_tipo": "FILA_PARTIDA_DESPLAZADA",
            }
        )

    return x, pd.DataFrame(reparaciones)


def detectar_solapamientos_periodos(resumen):
    """Detecta dos archivos que cubren días superpuestos de la misma cuenta lógica."""
    if resumen is None or resumen.empty:
        return pd.DataFrame()

    cols = [
        "empresa_uid", "empresa", "meta_codigo", "meta_nombre", "cuenta_logica_uid",
        "archivo", "periodo_inicio", "periodo_fin",
    ]
    b = resumen[cols].drop_duplicates().copy()
    b["periodo_inicio"] = pd.to_datetime(b["periodo_inicio"], errors="coerce")
    b["periodo_fin"] = pd.to_datetime(b["periodo_fin"], errors="coerce")
    b = b[b["periodo_inicio"].notna() & b["periodo_fin"].notna()]

    hallazgos = []
    for _, g in b.groupby("cuenta_logica_uid"):
        rows = list(g.to_dict("records"))
        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                a, c = rows[i], rows[j]
                if a["archivo"] == c["archivo"]:
                    continue
                ini = max(a["periodo_inicio"], c["periodo_inicio"])
                fin = min(a["periodo_fin"], c["periodo_fin"])
                if ini <= fin:
                    hallazgos.append(
                        {
                            "empresa": a["empresa"],
                            "meta_codigo": a["meta_codigo"],
                            "meta_nombre": a["meta_nombre"],
                            "archivo_1": a["archivo"],
                            "periodo_1": f"{a['periodo_inicio'].date()} → {a['periodo_fin'].date()}",
                            "archivo_2": c["archivo"],
                            "periodo_2": f"{c['periodo_inicio'].date()} → {c['periodo_fin'].date()}",
                            "inicio_solapamiento": ini.date(),
                            "fin_solapamiento": fin.date(),
                        }
                    )
    return pd.DataFrame(hallazgos)


def extraer_numero_folio(ref_norm):
    if not ref_norm:
        return None
    m = re.search(r"(\d+)$", str(ref_norm))
    return m.group(1) if m else None


def normalizar_referencia_base(ref):
    """
    Normaliza referencias de Clientes y Proveedores sin destruir identificadores.

    Tipos conciliables:
      - FOLIO_PREFIJO: H-123, NCTA-123, etc.
      - FOLIO_NUMERICO: 8729
      - DOCUMENTO_ALFANUMERICO: 824FEFE3, 0A11DE43, etc.

    Referencias descriptivas con espacios permanecen como OTRA_REFERENCIA.
    """
    if es_vacio(ref):
        return None, "VACIA", None

    if isinstance(ref, float) and ref.is_integer():
        s = str(int(ref))
    else:
        s = str(ref).strip()

    s = texto_norm(s)
    if s in REFERENCIAS_VACIAS:
        return None, "VACIA", None

    s = re.sub(
        r"^(?:FACTURA|FAC|FOLIO|REF|REFERENCIA)\s*[:.\-]?\s*",
        "",
        s,
    )

    prefijos = "|".join(sorted(PREFIJOS_FOLIO, key=len, reverse=True))
    m = re.fullmatch(rf"({prefijos})[\s\-_/.:]*(\d+)", s)
    if m:
        pref, num = m.groups()
        return f"{pref}{num}", "FOLIO_PREFIJO", pref

    if re.fullmatch(r"\d+", s):
        return s, "FOLIO_NUMERICO", None

    # IDs documentales frecuentes de proveedores (por ejemplo UUID corto/hex).
    # Se exige mezcla de letras y números y ausencia de espacios para evitar
    # convertir descripciones bancarias o conceptos libres en documentos.
    if (
        4 <= len(s) <= 40
        and re.fullmatch(r"[A-Z0-9]+", s)
        and re.search(r"[A-Z]", s)
        and re.search(r"\d", s)
    ):
        return s, "DOCUMENTO_ALFANUMERICO", None

    libre = re.sub(r"\s+", " ", s).strip()
    return libre if libre else None, "OTRA_REFERENCIA", None

def extraer_referencia_de_concepto(concepto):
    """
    Recupera un folio del concepto SOLO cuando existe un candidato documental
    inequívoco. No intenta inferir folios a partir de cualquier número.
    """
    s = texto_norm(concepto)
    if not s:
        return None

    prefijos = "|".join(sorted(PREFIJOS_FOLIO, key=len, reverse=True))
    patron = re.compile(
        rf"(?<![A-Z0-9])({prefijos})[\s\-_/.:]*(\d{{2,}})(?!\d)"
    )
    encontrados = {
        f"{m.group(1)}{m.group(2)}" for m in patron.finditer(s)
    }
    if len(encontrados) == 1:
        return next(iter(encontrados))
    return None


def enriquecer_referencias(df):
    df = df.copy()
    df["referencia_original"] = df["referencia"]

    refs_norm = []
    refs_tipo = []
    refs_prefijo = []
    refs_fuente = []
    refs_recuperadas = []

    for _, row in df.iterrows():
        original = row["referencia_original"]
        norm, tipo, prefijo = normalizar_referencia_base(original)
        recuperada = False
        fuente = "original"

        if norm is None:
            rec = extraer_referencia_de_concepto(row.get("concepto", ""))
            if rec:
                norm, tipo, prefijo = normalizar_referencia_base(rec)
                fuente = "concepto"
                recuperada = True
            else:
                fuente = "vacia"

        refs_norm.append(norm)
        refs_tipo.append(tipo)
        refs_prefijo.append(prefijo)
        refs_fuente.append(fuente)
        refs_recuperadas.append(recuperada)

    df["referencia_norm"] = refs_norm
    df["referencia_tipo"] = refs_tipo
    df["referencia_prefijo"] = refs_prefijo
    df["referencia_fuente"] = refs_fuente
    df["referencia_recuperada"] = refs_recuperadas
    df["referencia_numero"] = df["referencia_norm"].apply(extraer_numero_folio)

    df["tiene_referencia"] = df["referencia_norm"].notna()
    df["es_folio"] = df["referencia_tipo"].isin(
        ["FOLIO_PREFIJO", "FOLIO_NUMERICO"]
    )
    df["es_documento_conciliable"] = df["referencia_tipo"].isin(
        ["FOLIO_PREFIJO", "FOLIO_NUMERICO", "DOCUMENTO_ALFANUMERICO"]
    )
    return df

def cargar_archivo_robusto(file_bytes, file_name):
    """
    Lee Excel o CSV. CSV intenta UTF-8 antes de latin-1 para evitar mojibake.
    """
    bio = BytesIO(file_bytes)
    lower = file_name.lower()

    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(bio, header=None)

    # Incluso si la extensión es CSV, primero intentamos Excel por seguridad
    # cuando el contenido realmente lo sea.
    try:
        return pd.read_excel(bio, header=None)
    except Exception:
        pass

    errores = []
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            bio.seek(0)
            return pd.read_csv(
                bio,
                header=None,
                encoding=encoding,
                sep=None,
                engine="python",
            )
        except Exception as e:
            errores.append(f"{encoding}: {e}")

    raise ValueError(
        "No fue posible leer el archivo como Excel ni CSV. "
        + " | ".join(errores[-2:])
    )


def to_excel_workbook(tablas):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        usados = set()
        for nombre, df in tablas.items():
            hoja = re.sub(r"[\[\]\*\?/\\:]", "_", str(nombre))[:31] or "Datos"
            base = hoja
            i = 2
            while hoja in usados:
                suf = f"_{i}"
                hoja = (base[: 31 - len(suf)] + suf)
                i += 1
            usados.add(hoja)
            (df if df is not None else pd.DataFrame()).to_excel(
                writer, sheet_name=hoja, index=False
            )
    return output.getvalue()


def construir_empresa_uid(sistema_origen, empresa):
    """Identidad estable para impedir cruces entre sistemas o empresas distintas."""
    empresa_norm = texto_norm(empresa)
    if not empresa_norm:
        empresa_norm = "SIN_EMPRESA_IDENTIFICADA"
    return f"{sistema_origen}::{empresa_norm}"


def agregar_identidad_origen(df, sistema_origen, empresa, file_name):
    """Agrega identidad de sistema, empresa y cuenta a movimientos/resúmenes."""
    x = df.copy()
    empresa_txt = "" if empresa is None else str(empresa).strip()
    empresa_uid = construir_empresa_uid(sistema_origen, empresa_txt)
    x["sistema_origen"] = sistema_origen
    x["empresa"] = empresa_txt
    x["empresa_uid"] = empresa_uid
    x["cuenta_logica_uid"] = (
        empresa_uid + "::" + x["meta_codigo"].astype(str)
    )
    # cuenta_uid distingue la ocurrencia del mismo código dentro de cada archivo.
    x["cuenta_uid"] = (
        empresa_uid + "::" + str(file_name) + "::" + x["meta_codigo"].astype(str)
    )
    return x


# ==============================================================================
# 2. MOTOR ARPON DE LECTURA Y VALIDACIÓN
# ==============================================================================

def validar_formato_arpon(raw, file_name):
    """Valida que el archivo tenga la firma del Auxiliar de Cuentas de ARPON."""
    if raw.empty:
        raise ValueError(f"{file_name}: el archivo está vacío.")

    if raw.shape[1] < 7:
        raise ValueError(
            f"{file_name}: este auditor es exclusivo para ARPON y espera al menos "
            "7 columnas: Póliza | Fecha | Docto. | Concepto | Cargo | Abono | Saldo."
        )

    for _, row in raw.iterrows():
        vals = [texto_norm(row.iloc[i]) for i in range(7)]
        if (
            vals[0] == "POLIZA"
            and vals[1] == "FECHA"
            and vals[2].startswith("DOCTO")
            and vals[3] == "CONCEPTO"
            and vals[4] == "CARGO"
            and vals[5] == "ABONO"
            and vals[6] == "SALDO"
        ):
            return True

    primeras = raw.head(10).fillna("").astype(str).to_string(index=False, header=False)
    raise ValueError(
        f"{file_name}: el archivo no corresponde al Auxiliar de Cuentas de ARPON. "
        "Este auditor solo acepta exportaciones ARPON. Se espera la estructura "
        "Póliza | Fecha | Docto. | Concepto | Cargo | Abono | Saldo. "
        f"Primeras filas detectadas:\n{primeras[:800]}"
    )

def extraer_empresa_periodo_arpon(raw):
    """Extrae empresa y periodo del Auxiliar de Cuentas exportado por ARPON."""
    empresa = None
    periodo_inicio = pd.NaT
    periodo_fin = pd.NaT

    idx_aux = None
    texto_aux = None
    for idx in raw.index:
        for col in range(min(raw.shape[1], 4)):
            t = texto_norm(raw.iloc[idx, col])
            if "AUXILIAR DE CUENTAS" in t:
                idx_aux = idx
                texto_aux = str(raw.iloc[idx, col])
                break
        if idx_aux is not None:
            break

    if idx_aux is not None:
        for j in range(idx_aux - 1, -1, -1):
            candidato = raw.iloc[j, 0] if raw.shape[1] else None
            if es_vacio(candidato):
                continue
            c = str(candidato).strip()
            cn = texto_norm(c)
            if cn.startswith("MON ") or cn.startswith("TUE ") or cn.startswith("WED "):
                continue
            if cn.startswith("THU ") or cn.startswith("FRI ") or cn.startswith("SAT "):
                continue
            if cn.startswith("SUN "):
                continue
            empresa = c
            break

    if texto_aux:
        m = re.search(
            r"DEL\s+(.+?)\s+AL\s+(.+?)$",
            texto_norm(texto_aux),
            flags=re.I,
        )
        if m:
            periodo_inicio = parse_spanish_date(m.group(1).title())
            periodo_fin = parse_spanish_date(m.group(2).title())

    return empresa, periodo_inicio, periodo_fin


def validar_secuencia_saldo(movs, resumen):
    """
    Valida la ecuación física del auxiliar ARPON movimiento por movimiento:

        saldo_nuevo = saldo_anterior + cargo - abono

    La naturaleza de negocio (Clientes/Proveedores) se interpreta después y no
    altera esta validación de la fuente.
    """
    resultados = []
    mapa_si = resumen.set_index("cuenta_uid")["saldo_inicial"]
    mapa_sf = resumen.set_index("cuenta_uid")["saldo_final_aux"]

    for cuenta_uid, mm in movs.groupby("cuenta_uid"):
        mm = mm.sort_values("fila_origen").copy()
        saldo_ini = float(mapa_si.loc[cuenta_uid])
        saldo_final_reporte = float(mapa_sf.loc[cuenta_uid])

        prev = mm["saldo_acumulado"].shift(1)
        prev.iloc[0] = saldo_ini
        esperado = prev + mm["cargos"] - mm["abonos"]
        errores = (mm["saldo_acumulado"] - esperado).abs()

        ultimo = float(mm.iloc[-1]["saldo_acumulado"]) if len(mm) else saldo_ini
        error_final = abs(ultimo - saldo_final_reporte)

        resultados.append(
            {
                "cuenta_uid": cuenta_uid,
                "ecuacion_saldo_fuente": "SALDO_ANTERIOR + CARGO - ABONO",
                "n_errores_saldo_secuencia": int((errores > UMBRAL_TOLERANCIA).sum()),
                "max_error_saldo_secuencia": float(errores.max()) if len(errores) else 0.0,
                "ultimo_saldo_movimiento": ultimo,
                "error_ultimo_saldo_vs_total": float(error_final),
            }
        )

    return pd.DataFrame(resultados)

def procesar_formato_arpon(raw, file_name):
    """Procesa auxiliares ARPON de CLIENTES y PROVEEDORES de Hotel Quartz."""
    if raw.shape[1] < 7:
        raise ValueError(
            f"{file_name}: el auxiliar ARPON requiere al menos 7 columnas."
        )

    raw, reparaciones = reconstruir_filas_fragmentadas_arpon(raw, file_name)
    empresa, periodo_inicio, periodo_fin = extraer_empresa_periodo_arpon(raw)

    patron_header = re.compile(
        r"^CUENTA:\s*(\d+(?:-\d+){2,})\s*-\s*(.+)$",
        flags=re.I,
    )
    headers = []
    for idx in raw.index:
        t = texto_norm(raw.iloc[idx, 0])
        m = patron_header.match(t)
        if not m:
            continue
        saldo_ini = parse_amount(raw.iloc[idx, 6])
        if pd.isna(saldo_ini):
            raise ValueError(
                f"{file_name}: no pude leer el saldo inicial de la cuenta "
                f"{m.group(1)} en la fila Excel {idx + 1}."
            )
        nombre = m.group(2).strip()
        tipo_cuenta = detectar_tipo_cuenta(nombre)
        if tipo_cuenta == "NO_SOPORTADA":
            raise ValueError(
                f"{file_name}: la cuenta {m.group(1)} - {nombre} no es CLIENTES ni "
                "PROVEEDORES. Esta versión certifica únicamente esas dos familias."
            )
        headers.append(
            {
                "idx": idx,
                "codigo": m.group(1),
                "nombre": nombre,
                "tipo_cuenta": tipo_cuenta,
                "saldo_inicial": float(saldo_ini),
            }
        )

    if not headers:
        raise ValueError(
            f"{file_name}: no se detectó ninguna fila 'Cuenta: código - nombre'."
        )

    df = raw.copy()
    df["meta_codigo"] = pd.Series(index=df.index, dtype="object")
    df["meta_nombre"] = pd.Series(index=df.index, dtype="object")
    df["meta_tipo_cuenta"] = pd.Series(index=df.index, dtype="object")
    df["meta_saldo_inicial"] = pd.Series(index=df.index, dtype="float64")

    for h in headers:
        df.loc[h["idx"], "meta_codigo"] = h["codigo"]
        df.loc[h["idx"], "meta_nombre"] = h["nombre"]
        df.loc[h["idx"], "meta_tipo_cuenta"] = h["tipo_cuenta"]
        df.loc[h["idx"], "meta_saldo_inicial"] = h["saldo_inicial"]

    for c in ["meta_codigo", "meta_nombre", "meta_tipo_cuenta", "meta_saldo_inicial"]:
        df[c] = df[c].ffill()

    fechas_candidato = raw[1].apply(parse_spanish_date)
    etiquetas_b = raw[1].apply(texto_norm)
    is_mov = (
        fechas_candidato.notna()
        & raw[0].apply(lambda x: not es_vacio(x))
        & df["meta_codigo"].notna()
        & ~etiquetas_b.isin({"TOTALES", "NETO PERIODO"})
    )

    if not is_mov.any():
        raise ValueError(f"{file_name}: no se detectaron movimientos válidos.")

    movs = df[is_mov].copy().rename(
        columns={
            0: "poliza", 1: "fecha_raw", 2: "referencia", 3: "concepto",
            4: "cargos", 5: "abonos", 6: "saldo_acumulado",
        }
    )
    movs["tipo_poliza"] = (
        movs["poliza"].astype(str).str.extract(r"^([A-Za-z]+)", expand=False)
        .fillna("").str.upper()
    )
    movs["fila_origen"] = movs.index + 1
    movs["archivo"] = file_name
    movs["periodo_inicio"] = periodo_inicio
    movs["periodo_fin"] = periodo_fin
    movs["tipo_cuenta"] = movs["meta_tipo_cuenta"]
    movs = agregar_identidad_origen(movs, "ARPON", empresa or "", file_name)
    movs["fecha"] = movs["fecha_raw"].apply(parse_spanish_date)

    for c in ["cargos", "abonos", "saldo_acumulado"]:
        original = movs[c].copy()
        convertido = columna_a_monto(original)
        invalidos = convertido.isna()
        if invalidos.any():
            filas = movs.loc[invalidos, "fila_origen"].tolist()
            ejemplos = original[invalidos].astype(str).head(5).tolist()
            raise ValueError(
                f"{file_name}: valores no numéricos en '{c}' en "
                f"{len(filas)} movimiento(s). Filas {filas[:10]}; ejemplos: {ejemplos}."
            )
        movs[c] = convertido.astype(float)

    # Trazabilidad de las filas reconstruidas.
    if reparaciones.empty:
        movs["fila_reparada"] = False
        movs["fila_continuacion"] = np.nan
        movs["reparacion_tipo"] = ""
    else:
        mapa_cont = reparaciones.set_index("fila_principal")["fila_continuacion"].to_dict()
        mapa_tipo = reparaciones.set_index("fila_principal")["reparacion_tipo"].to_dict()
        movs["fila_reparada"] = movs["fila_origen"].isin(mapa_cont)
        movs["fila_continuacion"] = movs["fila_origen"].map(mapa_cont)
        movs["reparacion_tipo"] = movs["fila_origen"].map(mapa_tipo).fillna("")

    movs["concepto_norm"] = movs["concepto"].apply(concepto_norm)
    movs = enriquecer_referencias(movs)

    resumen_rows = []
    empresa_uid = construir_empresa_uid("ARPON", empresa or "")
    for h in headers:
        uid = f"{empresa_uid}::{file_name}::{h['codigo']}"
        mm = movs[movs["cuenta_uid"] == uid].sort_values("fila_origen")
        total_cargos = float(mm["cargos"].sum()) if len(mm) else 0.0
        total_abonos = float(mm["abonos"].sum()) if len(mm) else 0.0
        saldo_final = float(mm.iloc[-1]["saldo_acumulado"]) if len(mm) else h["saldo_inicial"]
        resumen_rows.append(
            {
                "archivo": file_name,
                "sistema_origen": "ARPON",
                "empresa": empresa or "",
                "empresa_uid": empresa_uid,
                "cuenta_logica_uid": f"{empresa_uid}::{h['codigo']}",
                "periodo_inicio": periodo_inicio,
                "periodo_fin": periodo_fin,
                "cuenta_uid": uid,
                "meta_codigo": h["codigo"],
                "meta_nombre": h["nombre"],
                "tipo_cuenta": h["tipo_cuenta"],
                "saldo_inicial": h["saldo_inicial"],
                "total_cargos": total_cargos,
                "total_abonos": total_abonos,
                "saldo_final_aux": saldo_final,
                "total_explicito_arpon": False,
            }
        )
    resumen = pd.DataFrame(resumen_rows)

    # Totales explícitos del reporte, asociados a la cuenta activa por ffill.
    mask_totales = raw[1].apply(texto_norm).eq("TOTALES")
    total_rows = raw.index[mask_totales].tolist()
    explicitos = []
    for idx in total_rows:
        codigo = df.loc[idx, "meta_codigo"]
        if pd.isna(codigo):
            continue
        tc, ta, sf = (parse_amount(raw.iloc[idx, c]) for c in (4, 5, 6))
        if any(pd.isna(v) for v in (tc, ta, sf)):
            raise ValueError(
                f"{file_name}: no fue posible leer la fila Totales (fila Excel {idx + 1})."
            )
        explicitos.append(
            {
                "meta_codigo": str(codigo),
                "total_cargos_exp": float(tc),
                "total_abonos_exp": float(ta),
                "saldo_final_exp": float(sf),
            }
        )

    if explicitos:
        exp = pd.DataFrame(explicitos).drop_duplicates("meta_codigo", keep="last")
        resumen = resumen.merge(exp, on="meta_codigo", how="left")
        tiene_exp = resumen["saldo_final_exp"].notna()
        resumen.loc[tiene_exp, "total_cargos"] = resumen.loc[tiene_exp, "total_cargos_exp"]
        resumen.loc[tiene_exp, "total_abonos"] = resumen.loc[tiene_exp, "total_abonos_exp"]
        resumen.loc[tiene_exp, "saldo_final_aux"] = resumen.loc[tiene_exp, "saldo_final_exp"]
        resumen.loc[tiene_exp, "total_explicito_arpon"] = True
        resumen = resumen.drop(
            columns=["total_cargos_exp", "total_abonos_exp", "saldo_final_exp"]
        )

    sum_mov = (
        movs.groupby("cuenta_uid", as_index=False)
        .agg(mov_cargos=("cargos", "sum"), mov_abonos=("abonos", "sum"))
    )
    resumen = resumen.merge(sum_mov, on="cuenta_uid", how="left")
    resumen[["mov_cargos", "mov_abonos"]] = resumen[["mov_cargos", "mov_abonos"]].fillna(0.0)
    resumen["dif_cargos_vs_total"] = resumen["total_cargos"] - resumen["mov_cargos"]
    resumen["dif_abonos_vs_total"] = resumen["total_abonos"] - resumen["mov_abonos"]

    # Solo una fila Totales explícita constituye una validación independiente.
    mal_detalle = resumen[
        resumen["total_explicito_arpon"]
        & (
            (resumen["dif_cargos_vs_total"].abs() > UMBRAL_TOLERANCIA)
            | (resumen["dif_abonos_vs_total"].abs() > UMBRAL_TOLERANCIA)
        )
    ]
    if not mal_detalle.empty:
        detalle = "; ".join(
            f"{r.meta_codigo}: Δcargo={r.dif_cargos_vs_total:,.2f}, "
            f"Δabono={r.dif_abonos_vs_total:,.2f}"
            for r in mal_detalle.itertuples()
        )
        raise ValueError(
            f"{file_name}: el detalle no amarra con la fila Totales del auxiliar ARPON. {detalle}"
        )

    sec = validar_secuencia_saldo(movs, resumen)
    resumen = resumen.merge(sec, on="cuenta_uid", how="left")
    errores_secuencia = int(resumen["n_errores_saldo_secuencia"].fillna(0).sum())
    max_error_secuencia = float(resumen["max_error_saldo_secuencia"].fillna(0).max())
    error_final_max = float(resumen["error_ultimo_saldo_vs_total"].fillna(0).max())

    if errores_secuencia:
        raise ValueError(
            f"{file_name}: se detectaron {errores_secuencia} movimiento(s) cuya secuencia "
            "de saldo acumulado no puede reproducirse con SALDO + CARGO - ABONO."
        )
    if error_final_max > UMBRAL_TOLERANCIA:
        raise ValueError(
            f"{file_name}: el último saldo de movimientos no coincide con el saldo final "
            f"ARPON. Diferencia máxima: ${error_final_max:,.2f}."
        )

    mask_neto = raw[1].apply(texto_norm).eq("NETO PERIODO")
    neto_periodo = None
    amarre_neto = None
    if mask_neto.any():
        idx_neto = raw.index[mask_neto][-1]
        candidatos = [raw.iloc[idx_neto, c] for c in range(4, min(7, raw.shape[1]))]
        for valor in candidatos:
            if es_vacio(valor):
                continue
            n = parse_amount(valor)
            if not pd.isna(n):
                neto_periodo = float(n)
                break
        if neto_periodo is not None:
            neto_calc = float(movs["cargos"].sum() - movs["abonos"].sum())
            amarre_neto = abs(abs(neto_calc) - abs(neto_periodo)) <= UMBRAL_TOLERANCIA

    todos_totales_exp = bool(len(resumen) and resumen["total_explicito_arpon"].all())
    gran_total_reportado = (
        float(resumen["saldo_final_aux"].sum()) if todos_totales_exp else None
    )
    gran_total_calculado = float(
        movs.groupby("cuenta_uid")["saldo_acumulado"].last().sum()
    )
    amarre_gran_total = (
        abs(gran_total_reportado - gran_total_calculado) <= UMBRAL_TOLERANCIA
        if gran_total_reportado is not None else None
    )
    amarre_totales = (
        bool(
            (
                (resumen.loc[resumen["total_explicito_arpon"], "dif_cargos_vs_total"].abs() <= UMBRAL_TOLERANCIA)
                & (resumen.loc[resumen["total_explicito_arpon"], "dif_abonos_vs_total"].abs() <= UMBRAL_TOLERANCIA)
            ).all()
        )
        if resumen["total_explicito_arpon"].any() else None
    )

    tipos = sorted(resumen["tipo_cuenta"].dropna().unique())
    diag = {
        "archivo": file_name,
        "sistema_origen": "ARPON",
        "formato": "ARPON_AUXILIAR_CUENTAS",
        "empresa": empresa or "",
        "tipo_cuenta": " | ".join(tipos),
        "periodo_inicio": periodo_inicio,
        "periodo_fin": periodo_fin,
        "n_headers": len(headers),
        "n_totales": len(total_rows),
        "n_movs": int(len(movs)),
        "n_filas_reconstruidas": int(len(reparaciones)),
        "gran_total_reportado": gran_total_reportado,
        "gran_total_calculado": gran_total_calculado,
        "gran_total": gran_total_reportado if gran_total_reportado is not None else gran_total_calculado,
        "origen_gran_total": "ARPON" if gran_total_reportado is not None else "CALCULADO",
        "suma_saldos_cuenta": float(resumen["saldo_final_aux"].sum()),
        "n_candidatos_no_header": 0,
        "amarre_gran_total": amarre_gran_total,
        "amarre_totales_detalle": amarre_totales,
        "neto_periodo": neto_periodo,
        "amarre_neto_periodo": amarre_neto,
        "n_errores_saldo_secuencia": errores_secuencia,
        "max_error_saldo_secuencia": max_error_secuencia,
        "max_error_ultimo_saldo_vs_total": error_final_max,
    }

    return movs.reset_index(drop=True), resumen.reset_index(drop=True), diag

def procesar_archivo_core(file_bytes, file_name):
    raw = cargar_archivo_robusto(file_bytes, file_name)
    validar_formato_arpon(raw, file_name)
    return procesar_formato_arpon(raw, file_name)


@st.cache_data(show_spinner=False)
def procesar_archivo_engine(file_bytes, file_name):
    return procesar_archivo_core(file_bytes, file_name)

# ==============================================================================
# 3. NATURALEZA CONTABLE Y CONCILIACIÓN ARPON
# ==============================================================================

def detectar_naturaleza(resumen, movs):
    """
    Separa dos conceptos que no deben confundirse:
      1) La ecuación física ARPON siempre es saldo + cargo - abono.
      2) La naturaleza de negocio depende de la familia de cuenta.

    CLIENTES    -> naturaleza DEUDORA, factor de pendiente +1.
    PROVEEDORES -> naturaleza ACREEDORA, factor de pendiente -1.
    """
    r = resumen.copy()
    r["esperado_arpon"] = r["saldo_inicial"] + r["total_cargos"] - r["total_abonos"]
    r["error_arpon"] = r["saldo_final_aux"] - r["esperado_arpon"]

    r["naturaleza"] = r["tipo_cuenta"].map(
        {"CLIENTES": "DEUDORA", "PROVEEDORES": "ACREEDORA"}
    ).fillna("INDETERMINADA")
    r["naturaleza_confianza"] = np.where(
        r["naturaleza"].eq("INDETERMINADA"), "BAJA", "ALTA"
    )
    r["factor_pendiente"] = r["tipo_cuenta"].apply(factor_pendiente_tipo)
    r["saldo_inicial_pendiente"] = r["saldo_inicial"] * r["factor_pendiente"]
    r["saldo_final_pendiente"] = r["saldo_final_aux"] * r["factor_pendiente"]
    return r

def aplicar_naturaleza_a_movimientos(movs, resumen_naturaleza):
    m = movs.copy()
    m["movimiento_id"] = np.arange(1, len(m) + 1)
    mapa = resumen_naturaleza.set_index("cuenta_uid")
    m["naturaleza"] = m["cuenta_uid"].map(mapa["naturaleza"])
    m["factor_pendiente"] = m["cuenta_uid"].map(mapa["factor_pendiente"])

    # Cambio físico en ARPON y cambio normalizado del saldo pendiente.
    m["efecto_arpon"] = m["cargos"] - m["abonos"]
    m["efecto_natural"] = m["efecto_arpon"] * m["factor_pendiente"]
    m["saldo_pendiente"] = m["saldo_acumulado"] * m["factor_pendiente"]
    m["importe_abs"] = m["efecto_natural"].abs().round(2)
    return m

def marcar_duplicados_exactos(movs):
    m = movs.copy()
    subset = [
        "cuenta_logica_uid", "fecha", "tipo_poliza", "poliza", "concepto_norm",
        "referencia_norm", "cargos", "abonos"
    ]
    m["posible_duplicado_exacto"] = m.duplicated(subset=subset, keep=False)
    return m

def analizar_saldos(movs, resumen_naturaleza):
    """Conciliación por saldo pendiente normalizado para Clientes y Proveedores."""
    m = movs.copy()
    r = resumen_naturaleza.copy()

    con_ref = m[m["tiene_referencia"]].groupby("cuenta_uid")["efecto_natural"].sum(min_count=1)
    sin_ref = m[~m["tiene_referencia"]].groupby("cuenta_uid")["efecto_natural"].sum(min_count=1)

    sin_ref_stats = (
        m[~m["tiene_referencia"]].groupby("cuenta_uid").agg(
            n_sin_referencia=("cuenta_uid", "size"),
            cargos_sin_referencia=("cargos", "sum"),
            abonos_sin_referencia=("abonos", "sum"),
        )
    )
    rec_stats = m[m["referencia_recuperada"]].groupby("cuenta_uid").size().rename("n_refs_recuperadas")
    otras_ref_stats = m[m["referencia_tipo"].eq("OTRA_REFERENCIA")].groupby("cuenta_uid").size().rename("n_referencias_libres")
    neg_stats = m[(m["cargos"] < 0) | (m["abonos"] < 0)].groupby("cuenta_uid").size().rename("n_montos_negativos")
    dup_stats = m[m["posible_duplicado_exacto"]].groupby("cuenta_uid").size().rename("n_filas_posible_duplicado")
    repar_stats = m[m["fila_reparada"]].groupby("cuenta_uid").size().rename("n_filas_reconstruidas")

    r["movs_con_referencia"] = r["cuenta_uid"].map(con_ref).fillna(0.0)
    r["movs_sin_referencia"] = r["cuenta_uid"].map(sin_ref).fillna(0.0)
    r = r.merge(sin_ref_stats, left_on="cuenta_uid", right_index=True, how="left")
    for c in ["n_sin_referencia", "cargos_sin_referencia", "abonos_sin_referencia"]:
        r[c] = r[c].fillna(0)

    r["importe_bruto_sin_referencia"] = r["cargos_sin_referencia"].abs() + r["abonos_sin_referencia"].abs()
    r["n_refs_recuperadas"] = r["cuenta_uid"].map(rec_stats).fillna(0).astype(int)
    r["n_referencias_libres"] = r["cuenta_uid"].map(otras_ref_stats).fillna(0).astype(int)
    r["n_montos_negativos"] = r["cuenta_uid"].map(neg_stats).fillna(0).astype(int)
    r["n_filas_posible_duplicado"] = r["cuenta_uid"].map(dup_stats).fillna(0).astype(int)
    r["n_filas_reconstruidas"] = r["cuenta_uid"].map(repar_stats).fillna(0).astype(int)

    r["saldo_esperado_motor"] = r["saldo_inicial_pendiente"] + r["movs_con_referencia"] + r["movs_sin_referencia"]
    r["descuadre_origen"] = r["saldo_final_pendiente"] - r["saldo_esperado_motor"]
    r["cuadra"] = r["descuadre_origen"].abs() <= UMBRAL_TOLERANCIA
    r["tiene_arrastre"] = r["saldo_inicial_pendiente"].abs() > UMBRAL_TOLERANCIA
    r["tiene_sin_referencia"] = r["n_sin_referencia"] > 0
    r["tiene_montos_negativos"] = r["n_montos_negativos"] > 0

    def estado(row):
        if row["naturaleza"] == "INDETERMINADA":
            return "⚫ Naturaleza indeterminada"
        if abs(row["descuadre_origen"]) > UMBRAL_TOLERANCIA:
            return "🟠 Total fuente ≠ Detalle"
        if row["n_sin_referencia"] > 0:
            return "🔴 Movimientos sin referencia"
        if row["n_montos_negativos"] > 0:
            return "🟣 Montos negativos / reversos"
        return "🟢 OK"

    r["estado"] = r.apply(estado, axis=1)
    return r

# ==============================================================================
# 4. FOLIOS, REFERENCIAS Y CRUCES ARPON
# ==============================================================================

def analizar_folios(movs, fecha_corte):
    """Analiza documentos conciliables abiertos a través de todos los periodos cargados."""
    mv = movs[
        movs["es_documento_conciliable"]
        & movs["efecto_natural"].notna()
    ].copy()

    if mv.empty:
        return pd.DataFrame(
            columns=[
                "sistema_origen", "empresa", "tipo_cuenta", "archivo", "archivos",
                "meta_codigo", "meta_nombre", "naturaleza", "referencia_norm",
                "primera_fecha", "ultima_fecha", "n_movs", "cargos", "abonos",
                "saldo_natural", "dias", "antiguedad_observada", "tipo_saldo",
                "multiples_movimientos", "posible_duplicado_exacto"
            ]
        )

    g = (
        mv.groupby(
            [
                "sistema_origen", "empresa_uid", "empresa", "tipo_cuenta",
                "cuenta_logica_uid", "meta_codigo", "meta_nombre",
                "naturaleza", "referencia_norm"
            ],
            as_index=False,
        )
        .agg(
            primera_fecha=("fecha", "min"),
            ultima_fecha=("fecha", "max"),
            n_movs=("efecto_natural", "size"),
            cargos=("cargos", "sum"),
            abonos=("abonos", "sum"),
            saldo_natural=("efecto_natural", "sum"),
            posible_duplicado_exacto=("posible_duplicado_exacto", "max"),
            archivos=("archivo", lambda x: " | ".join(sorted(set(map(str, x))))),
        )
    )
    g["archivo"] = g["archivos"]

    vivos = g[g["saldo_natural"].abs() > UMBRAL_FOLIO].copy()
    corte = pd.Timestamp(fecha_corte)
    vivos["dias"] = (corte - vivos["primera_fecha"]).dt.days

    def bucket(d):
        if pd.isna(d): return "sin fecha"
        if d < 0: return "fecha posterior al corte"
        if d <= 30: return "0-30"
        if d <= 60: return "31-60"
        if d <= 90: return "61-90"
        return "90+"

    vivos["antiguedad_observada"] = vivos["dias"].apply(bucket)

    def tipo_saldo(row):
        s = row["saldo_natural"]
        if row["tipo_cuenta"] == "CLIENTES":
            return "🔵 Pendiente de cobro" if s > 0 else "🔴 Saldo contrario a naturaleza / a favor"
        if row["tipo_cuenta"] == "PROVEEDORES":
            return "🟣 Pendiente de pago" if s > 0 else "🔴 Saldo contrario a naturaleza / a favor"
        return "⚫ Naturaleza indeterminada"

    vivos["tipo_saldo"] = vivos.apply(tipo_saldo, axis=1)
    vivos["multiples_movimientos"] = vivos["n_movs"] > 2
    return vivos.sort_values(["dias", "saldo_natural"], ascending=[False, False])

def detectar_cruces_por_referencia(movs):
    """Busca el mismo documento en cuentas distintas del mismo tipo y empresa."""
    mv = movs[
        movs["es_documento_conciliable"]
        & movs["efecto_natural"].notna()
    ].copy()
    if mv.empty:
        return pd.DataFrame()

    por_cuenta = (
        mv.groupby(
            [
                "sistema_origen", "empresa_uid", "empresa", "tipo_cuenta",
                "referencia_norm", "cuenta_logica_uid", "meta_codigo", "meta_nombre",
                "naturaleza"
            ],
            as_index=False,
        )
        .agg(
            cargos=("cargos", "sum"), abonos=("abonos", "sum"),
            efecto_natural=("efecto_natural", "sum"), n_movs=("efecto_natural", "size"),
            archivos=("archivo", lambda x: " | ".join(sorted(set(map(str, x))))),
        )
    )

    claves_ref = ["sistema_origen", "empresa_uid", "tipo_cuenta", "referencia_norm"]
    nivel_ref = (
        por_cuenta.groupby(claves_ref)
        .agg(
            num_cuentas=("cuenta_logica_uid", "nunique"),
            hay_positivo=("efecto_natural", lambda x: (x > UMBRAL_FOLIO).any()),
            hay_negativo=("efecto_natural", lambda x: (x < -UMBRAL_FOLIO).any()),
            neto_global=("efecto_natural", "sum"),
        ).reset_index()
    )
    refs = nivel_ref[
        (nivel_ref["num_cuentas"] > 1) & nivel_ref["hay_positivo"] & nivel_ref["hay_negativo"]
    ].copy()
    if refs.empty:
        return pd.DataFrame()

    detalle = por_cuenta.merge(
        refs[claves_ref + ["num_cuentas", "neto_global"]], on=claves_ref, how="inner"
    )
    detalle["amarre_aprox"] = detalle["neto_global"].abs() <= UMBRAL_TOLERANCIA
    detalle["nivel_evidencia"] = np.where(
        detalle["amarre_aprox"], "ALTA - neto a cero", "MEDIA - efectos opuestos con remanente"
    )
    return detalle.sort_values(
        ["sistema_origen", "empresa_uid", "tipo_cuenta", "referencia_norm", "efecto_natural"],
        ascending=[True, True, True, True, False],
    )

def detectar_coincidencias_por_evidencia(movs):
    """Coincidencias exactas por fecha + concepto + importe dentro del mismo tipo de cuenta."""
    mv = movs[
        movs["efecto_natural"].notna()
        & (movs["importe_abs"] > UMBRAL_FOLIO)
        & movs["concepto_norm"].ne("")
    ].copy()
    if mv.empty:
        return pd.DataFrame()

    mv = mv[mv["efecto_natural"].abs() > UMBRAL_FOLIO].copy()
    claves = [
        "sistema_origen", "empresa_uid", "tipo_cuenta", "fecha", "concepto_norm", "importe_abs"
    ]
    grupos = (
        mv.groupby(claves)
        .agg(
            num_cuentas=("cuenta_logica_uid", "nunique"),
            hay_positivo=("efecto_natural", lambda x: (x > 0).any()),
            hay_negativo=("efecto_natural", lambda x: (x < 0).any()),
            n_movs_grupo=("efecto_natural", "size"),
            n_positivos=("efecto_natural", lambda x: int((x > 0).sum())),
            n_negativos=("efecto_natural", lambda x: int((x < 0).sum())),
            neto_grupo=("efecto_natural", "sum"),
        ).reset_index()
    )
    validos = grupos[
        (grupos["num_cuentas"] > 1) & grupos["hay_positivo"] & grupos["hay_negativo"]
    ].copy()
    if validos.empty:
        return pd.DataFrame()

    validos["amarre_aprox"] = validos["neto_grupo"].abs() <= UMBRAL_TOLERANCIA
    validos["nivel_evidencia"] = np.select(
        [
            validos["amarre_aprox"] & validos["n_positivos"].eq(1) & validos["n_negativos"].eq(1),
            validos["amarre_aprox"],
        ],
        ["ALTA - correspondencia 1:1", "MEDIA - neto cero con múltiples movimientos"],
        default="BAJA - coincidencia parcial con remanente",
    )
    validos["evidencia_id"] = np.arange(1, len(validos) + 1)
    det = mv.merge(validos, on=claves, how="inner")
    cols = [
        "evidencia_id", "nivel_evidencia", "amarre_aprox", "movimiento_id", "fila_origen",
        "sistema_origen", "empresa_uid", "empresa", "tipo_cuenta", "fecha", "concepto",
        "concepto_norm", "importe_abs", "archivo", "cuenta_uid", "cuenta_logica_uid",
        "meta_codigo", "meta_nombre", "naturaleza", "referencia_original", "referencia_norm",
        "referencia_fuente", "cargos", "abonos", "efecto_natural", "num_cuentas",
        "n_movs_grupo", "n_positivos", "n_negativos", "neto_grupo"
    ]
    return det[cols].sort_values(["evidencia_id", "efecto_natural"], ascending=[True, False])

def marcar_movimientos_conciliacion(movs, cruces_ref, evidencias):
    """Marca conciliaciones sin alterar ni eliminar movimientos de origen."""
    m = movs.copy()
    if "movimiento_id" not in m.columns:
        m["movimiento_id"] = np.arange(1, len(m) + 1)

    registros = {
        int(mid): {"estado": "SIN MARCA", "nivel": "", "criterios": [], "codigos": []}
        for mid in m["movimiento_id"]
    }
    rango_estado = {"SIN MARCA": 0, "REVISAR": 1, "CONCILIADO": 2}
    rango_nivel = {"": 0, "BAJA": 1, "MEDIA": 2, "ALTA": 3}

    def registrar(ids, estado, nivel, criterio, codigo):
        nivel_base = str(nivel).split(" - ", 1)[0].strip().upper()
        if nivel_base not in rango_nivel:
            nivel_base = "BAJA"
        for mid in ids:
            reg = registros.get(int(mid))
            if reg is None:
                continue
            if rango_estado[estado] > rango_estado[reg["estado"]]: reg["estado"] = estado
            if rango_nivel[nivel_base] > rango_nivel[reg["nivel"]]: reg["nivel"] = nivel_base
            if criterio not in reg["criterios"]: reg["criterios"].append(criterio)
            if codigo not in reg["codigos"]: reg["codigos"].append(codigo)

    docs = m[
        m["es_documento_conciliable"]
        & m["efecto_natural"].notna()
        & (m["efecto_natural"].abs() > UMBRAL_FOLIO)
    ].copy()
    if not docs.empty:
        claves_cuenta = [
            "sistema_origen", "empresa_uid", "tipo_cuenta", "cuenta_logica_uid", "referencia_norm"
        ]
        grupos_cuenta = (
            docs.groupby(claves_cuenta, as_index=False)
            .agg(
                n_movs=("movimiento_id", "size"),
                hay_positivo=("efecto_natural", lambda x: (x > UMBRAL_FOLIO).any()),
                hay_negativo=("efecto_natural", lambda x: (x < -UMBRAL_FOLIO).any()),
                neto_cuenta=("efecto_natural", "sum"),
            )
        )
        grupos_cuenta = grupos_cuenta[
            (grupos_cuenta["n_movs"] > 1) & grupos_cuenta["hay_positivo"] & grupos_cuenta["hay_negativo"]
        ]
        for _, grupo in grupos_cuenta.iterrows():
            amarra = abs(float(grupo["neto_cuenta"])) <= UMBRAL_TOLERANCIA
            mask = (
                m["sistema_origen"].eq(grupo["sistema_origen"])
                & m["empresa_uid"].eq(grupo["empresa_uid"])
                & m["tipo_cuenta"].eq(grupo["tipo_cuenta"])
                & m["cuenta_logica_uid"].eq(grupo["cuenta_logica_uid"])
                & m["referencia_norm"].eq(grupo["referencia_norm"])
                & (m["efecto_natural"].abs() > UMBRAL_FOLIO)
            )
            registrar(
                m.loc[mask, "movimiento_id"],
                "CONCILIADO" if amarra else "REVISAR",
                "ALTA" if amarra else "MEDIA",
                "DOCUMENTO SALDADO" if amarra else "DOCUMENTO PARCIAL",
                f"DOC:{grupo['referencia_norm']}",
            )

    if cruces_ref is not None and not cruces_ref.empty:
        claves = [
            "sistema_origen", "empresa_uid", "tipo_cuenta", "referencia_norm",
            "amarre_aprox", "nivel_evidencia",
        ]
        for _, grupo in cruces_ref[claves].drop_duplicates().iterrows():
            mask = (
                m["sistema_origen"].eq(grupo["sistema_origen"])
                & m["empresa_uid"].eq(grupo["empresa_uid"])
                & m["tipo_cuenta"].eq(grupo["tipo_cuenta"])
                & m["referencia_norm"].eq(grupo["referencia_norm"])
                & m["es_documento_conciliable"]
                & (m["efecto_natural"].abs() > UMBRAL_FOLIO)
            )
            registrar(
                m.loc[mask, "movimiento_id"],
                "CONCILIADO" if bool(grupo["amarre_aprox"]) else "REVISAR",
                grupo["nivel_evidencia"], "DOCUMENTO ENTRE CUENTAS",
                f"REF:{grupo['referencia_norm']}",
            )

    if evidencias is not None and not evidencias.empty:
        grupos_evidencia = evidencias[["evidencia_id", "nivel_evidencia", "amarre_aprox"]].drop_duplicates()
        for _, grupo in grupos_evidencia.iterrows():
            ids = evidencias.loc[evidencias["evidencia_id"].eq(grupo["evidencia_id"]), "movimiento_id"].drop_duplicates()
            nivel = str(grupo["nivel_evidencia"]).split(" - ", 1)[0].upper()
            criterio = "EVIDENCIA 1:1" if nivel == "ALTA" else "EVIDENCIA GRUPAL" if nivel == "MEDIA" else "COINCIDENCIA PARCIAL"
            # Solo correspondencia 1:1 exacta se eleva automáticamente a CONCILIADO.
            estado = "CONCILIADO" if bool(grupo["amarre_aprox"]) and nivel == "ALTA" else "REVISAR"
            registrar(ids, estado, grupo["nivel_evidencia"], criterio, f"EVD:{int(grupo['evidencia_id'])}")

    m["conciliacion_estado"] = m["movimiento_id"].map(lambda mid: registros[int(mid)]["estado"])
    m["conciliacion_nivel"] = m["movimiento_id"].map(lambda mid: registros[int(mid)]["nivel"])
    m["conciliacion_criterio"] = m["movimiento_id"].map(lambda mid: " + ".join(registros[int(mid)]["criterios"]))
    m["conciliacion_codigo"] = m["movimiento_id"].map(lambda mid: " | ".join(registros[int(mid)]["codigos"]))
    m["conciliacion_marcada"] = m["conciliacion_estado"].ne("SIN MARCA")
    return m

def _buscar_fila_encabezado_arpon(ws):
    limite = min(ws.max_row, 60)
    for fila in range(1, limite + 1):
        vals = [texto_norm(ws.cell(fila, col).value) for col in range(1, 8)]
        if (
            vals[0] == "POLIZA"
            and vals[1] == "FECHA"
            and vals[2].startswith("DOCTO")
            and vals[3] == "CONCEPTO"
            and vals[4] == "CARGO"
            and vals[5] == "ABONO"
            and vals[6] == "SALDO"
        ):
            return fila
    raise ValueError("No se encontró el encabezado ARPON en el libro de origen.")


def _libro_desde_archivo(file_bytes, file_name):
    lower = file_name.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        return load_workbook(
            BytesIO(file_bytes),
            keep_vba=lower.endswith(".xlsm"),
            keep_links=True,
        )

    # CSV/XLS se convierten a XLSX para poder entregar el marcado visual.
    raw = cargar_archivo_robusto(file_bytes, file_name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Auxiliar ARPON"
    for fila_idx, valores in enumerate(
        raw.itertuples(index=False, name=None), start=1
    ):
        for col_idx, valor in enumerate(valores, start=1):
            if pd.isna(valor):
                valor = None
            elif isinstance(valor, pd.Timestamp):
                valor = valor.to_pydatetime()
            ws.cell(fila_idx, col_idx, valor)
    return wb


def construir_auxiliar_marcado(file_bytes, file_name, marcas):
    """Conserva el auxiliar original y agrega marcas de conciliación/reconstrucción."""
    wb = _libro_desde_archivo(file_bytes, file_name)
    ws = wb.worksheets[0]
    fila_header = _buscar_fila_encabezado_arpon(ws)

    columna_estado = None
    for col in range(8, ws.max_column + 1):
        if texto_norm(ws.cell(fila_header, col).value) == "CONCILIACION":
            columna_estado = col
            break
    if columna_estado is None:
        columna_estado = max(8, ws.max_column + 1)

    celda_header = ws.cell(fila_header, columna_estado)
    fuente_header = ws.cell(fila_header, 7)
    if celda_header.coordinate != fuente_header.coordinate:
        celda_header._style = copy(fuente_header._style)
        celda_header.number_format = fuente_header.number_format
        celda_header.alignment = copy(fuente_header.alignment)
    celda_header.value = "Conciliación / Auditoría"
    celda_header.font = Font(
        name=celda_header.font.name, size=celda_header.font.size,
        bold=True, color=celda_header.font.color,
    )
    celda_header.alignment = Alignment(horizontal="center", vertical="center")

    verde = PatternFill("solid", fgColor="C6EFCE")
    verde_suave = PatternFill("solid", fgColor="EAF4E3")
    amarillo = PatternFill("solid", fgColor="FFEB9C")
    amarillo_suave = PatternFill("solid", fgColor="FFF7D6")
    azul = PatternFill("solid", fgColor="D9EAF7")
    azul_suave = PatternFill("solid", fgColor="EEF6FC")

    leyenda = [
        (1, "Marca de auditoría · ℹ Azul = fila ARPON reconstruida"),
        (2, "✓ Verde = conciliado"),
        (3, "⚠ Amarillo = revisar remanente"),
    ]
    for fila, texto_leyenda in leyenda:
        celda = ws.cell(fila, columna_estado)
        if es_vacio(celda.value):
            celda.value = texto_leyenda
            celda.font = Font(bold=(fila == 1), color="1F1F1F", size=10)
            celda.fill = verde if fila == 2 else amarillo if fila == 3 else azul if fila == 4 else verde_suave

    for _, marca in marcas.sort_values("fila_origen").iterrows():
        fila = int(marca["fila_origen"])
        if fila < 1 or fila > ws.max_row:
            continue

        tiene_conc = bool(marca.get("conciliacion_marcada", False))
        reparada = bool(marca.get("fila_reparada", False))

        if tiene_conc:
            conciliado = marca["conciliacion_estado"] == "CONCILIADO"
            simbolo = "✓" if conciliado else "⚠"
            texto_marca = (
                f"{simbolo} {marca['conciliacion_estado']} · "
                f"{marca['conciliacion_criterio']} · {marca['conciliacion_codigo']}"
            )
            if reparada:
                texto_marca += f" · ℹ reconstruida con fila {int(marca['fila_continuacion'])}"
            fill = verde if conciliado else amarillo
            fill_row = verde_suave if conciliado else amarillo_suave
            font_color = "006100" if conciliado else "9C6500"
        else:
            texto_marca = f"ℹ RECONSTRUIDA PARA AUDITORÍA · continuación fila {int(marca['fila_continuacion'])}"
            fill = azul
            fill_row = azul_suave
            font_color = "1F4E78"

        celda = ws.cell(fila, columna_estado)
        celda.value = texto_marca
        celda.fill = fill
        celda.font = Font(bold=True, color=font_color, size=10)
        celda.alignment = Alignment(vertical="center", wrap_text=False)
        for col in range(1, columna_estado):
            origen = ws.cell(fila, col)
            if origen.fill is None or origen.fill.fill_type is None:
                origen.fill = fill_row

        if reparada and pd.notna(marca.get("fila_continuacion")):
            fcont = int(marca["fila_continuacion"])
            if 1 <= fcont <= ws.max_row:
                ccont = ws.cell(fcont, columna_estado)
                ccont.value = f"↳ Continuación usada para reconstruir fila {fila}"
                ccont.fill = azul
                ccont.font = Font(color="1F4E78", italic=True, size=10)

    letra_estado = ws.cell(1, columna_estado).column_letter
    ws.column_dimensions[letra_estado].width = max(ws.column_dimensions[letra_estado].width or 0, 64)

    output = BytesIO()
    extension = ".xlsm" if file_name.lower().endswith(".xlsm") else ".xlsx"
    wb.save(output)
    nombre_salida = f"{Path(file_name).stem}_MARCADO{extension}"
    return output.getvalue(), nombre_salida

@st.cache_data(show_spinner=False)
def construir_descarga_auxiliares_marcados(archivos, movs):
    """Devuelve XLSX/XLSM o ZIP incluyendo conciliaciones y filas reconstruidas."""
    resultados = []
    nombres_usados = set()

    for file_name, file_bytes in archivos:
        marcas = movs[
            movs["archivo"].eq(file_name)
            & (movs["conciliacion_marcada"] | movs["fila_reparada"])
        ].copy()
        data, nombre = construir_auxiliar_marcado(file_bytes, file_name, marcas)

        base = Path(nombre).stem
        extension = Path(nombre).suffix
        candidato = nombre
        i = 2
        while candidato in nombres_usados:
            candidato = f"{base}_{i}{extension}"
            i += 1
        nombres_usados.add(candidato)
        resultados.append((candidato, data))

    if len(resultados) == 1:
        nombre, data = resultados[0]
        mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if nombre.lower().endswith(".xlsm")
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        return data, nombre, mime

    output = BytesIO()
    with ZipFile(output, mode="w", compression=ZIP_DEFLATED) as zf:
        for nombre, data in resultados:
            zf.writestr(nombre, data)
    return output.getvalue(), "auxiliares_ARPON_MARCADOS.zip", "application/zip"

def tabla_referencias(movs):
    cols = [
        "sistema_origen", "empresa", "tipo_cuenta", "archivo", "fila_origen",
        "fila_reparada", "fila_continuacion", "fecha", "meta_codigo", "meta_nombre",
        "concepto", "referencia_original", "referencia_norm", "referencia_tipo",
        "es_documento_conciliable", "referencia_fuente", "referencia_recuperada",
        "cargos", "abonos", "naturaleza", "efecto_natural",
        "conciliacion_estado", "conciliacion_nivel", "conciliacion_criterio", "conciliacion_codigo",
    ]
    return movs[cols].copy()

def aplicar_filtros_tabla(
    df,
    empresas=None,
    archivos=None,
    cuentas=None,
    fecha_desde=None,
    fecha_hasta=None,
    busqueda="",
):
    """Aplica una misma barra de filtros a tablas con estructuras distintas."""
    if df is None:
        return pd.DataFrame()

    x = df.copy()
    if empresas is not None and "empresa" in x.columns:
        x = x[x["empresa"].fillna("").astype(str).isin(empresas)]

    if archivos is not None:
        if "archivos" in x.columns:
            if not archivos:
                x = x.iloc[0:0]
            else:
                patron = "|".join(re.escape(str(a)) for a in archivos)
                x = x[
                    x["archivos"].fillna("").astype(str)
                    .str.contains(patron, regex=True)
                ]
        elif "archivo" in x.columns:
            x = x[x["archivo"].fillna("").astype(str).isin(archivos)]

    if cuentas is not None and "meta_codigo" in x.columns:
        x = x[x["meta_codigo"].fillna("").astype(str).isin(cuentas)]

    columna_fecha = None
    for candidata in ["fecha", "primera_fecha"]:
        if candidata in x.columns:
            columna_fecha = candidata
            break
    if columna_fecha and fecha_desde is not None and fecha_hasta is not None:
        fechas = pd.to_datetime(x[columna_fecha], errors="coerce")
        desde = pd.Timestamp(fecha_desde)
        hasta = pd.Timestamp(fecha_hasta) + pd.Timedelta(days=1)
        x = x[fechas.ge(desde) & fechas.lt(hasta)]

    termino = texto_norm(busqueda)
    if termino:
        columnas_busqueda = [
            c for c in [
                "archivo", "archivos", "meta_codigo", "meta_nombre",
                "poliza", "tipo_poliza", "concepto", "concepto_norm",
                "referencia_original", "referencia_norm", "estado",
                "naturaleza", "nivel_evidencia", "conciliacion_estado",
                "conciliacion_criterio", "conciliacion_codigo",
            ]
            if c in x.columns
        ]
        if columnas_busqueda:
            coincide = pd.Series(False, index=x.index)
            for col in columnas_busqueda:
                coincide = coincide | x[col].apply(texto_norm).str.contains(
                    termino, regex=False, na=False
                )
            x = x[coincide]

    return x


# ==============================================================================
# 5. UI / UX · v5.3 MINIMAL FINANCE
# ==============================================================================


def _inyectar_estilos_ui():
    st.markdown(
        """
        <style>
        :root {
            --bg: #FFFFFF;
            --surface: #FFFFFF;
            --surface-soft: #F8FAFC;
            --text: #111827;
            --muted: #667085;
            --border: #E5E7EB;
            --border-strong: #D0D5DD;
            --accent: #928B57;
            --success: #067647;
            --warning: #B54708;
            --danger: #B42318;
            --info: #475467;
        }

        .stApp { background: var(--bg); }
        [data-testid="stHeader"] {
            background: rgba(255,255,255,.95);
            border-bottom: 1px solid #F2F4F7;
            backdrop-filter: blur(8px);
        }
        [data-testid="stSidebar"] { display:none; }

        .block-container {
            max-width: 1360px;
            padding-top: 1.15rem;
            padding-bottom: 3.5rem;
        }

        html, body, [class*="css"] { color: var(--text); }
        h1, h2, h3, h4 { color: var(--text); letter-spacing: -.025em; }
        .stCaptionContainer p, p { color: var(--muted); }

        .topbar {
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:20px;
            padding: 3px 0 18px;
            border-bottom:1px solid var(--border);
            margin-bottom:22px;
        }
        .brand {
            display:flex;
            align-items:center;
            gap:12px;
            min-width:0;
        }
        .brand-mark {
            width:10px;
            height:34px;
            border-radius:999px;
            background:var(--accent);
            flex:0 0 auto;
        }
        .brand-title {
            font-size:1.03rem;
            font-weight:760;
            letter-spacing:-.02em;
            color:var(--text);
            line-height:1.1;
        }
        .brand-sub {
            color:var(--muted);
            font-size:.76rem;
            margin-top:4px;
        }
        .version {
            color:#98A2B3;
            font-size:.73rem;
            white-space:nowrap;
        }

        .page-head { margin: 4px 0 18px; }
        .page-head h1 {
            font-size: clamp(1.55rem, 2.5vw, 2.05rem);
            font-weight:760;
            margin:0 0 5px;
        }
        .page-head p { margin:0; color:var(--muted); font-size:.92rem; }

        .section-head { margin: 26px 0 12px; }
        .section-head h2 { margin:0 0 3px; font-size:1.05rem; font-weight:740; }
        .section-head p { margin:0; color:var(--muted); font-size:.84rem; }

        .status-line {
            display:flex;
            align-items:center;
            gap:9px;
            padding:10px 12px;
            border:1px solid var(--border);
            border-radius:10px;
            background:#fff;
            margin:0 0 16px;
            font-size:.84rem;
            color:var(--muted);
        }
        .status-line strong { color:var(--text); font-weight:700; }
        .status-dot { width:7px; height:7px; border-radius:999px; flex:0 0 auto; }
        .status-line.success .status-dot { background:var(--success); }
        .status-line.warning .status-dot { background:var(--warning); }
        .status-line.danger .status-dot { background:var(--danger); }
        .status-line.info .status-dot { background:var(--info); }

        .meta-row {
            display:flex;
            flex-wrap:wrap;
            gap:7px;
            margin-top:12px;
        }
        .meta-chip {
            display:inline-flex;
            align-items:center;
            min-height:28px;
            padding:4px 9px;
            border-radius:7px;
            border:1px solid var(--border);
            background:#fff;
            color:#475467;
            font-size:.74rem;
            font-weight:620;
        }

        /* Streamlit containers become restrained utility panels. */
        [data-testid="stVerticalBlockBorderWrapper"] {
            border-color: var(--border) !important;
            border-radius: 12px !important;
            box-shadow: none !important;
        }
        [data-testid="stExpander"] {
            border:1px solid var(--border) !important;
            border-radius:10px !important;
            background:#fff !important;
        }
        [data-testid="stExpander"] summary { font-weight:680; color:var(--text); }

        /* Metrics: no card color; just clean financial figures. */
        div[data-testid="stMetric"] {
            background:#fff;
            border:1px solid var(--border);
            border-radius:10px;
            padding:14px 15px 13px;
            box-shadow:none;
        }
        div[data-testid="stMetricLabel"] { color:var(--muted); font-weight:650; font-size:.78rem; }
        div[data-testid="stMetricValue"] { color:var(--text); font-weight:760; letter-spacing:-.035em; }
        div[data-testid="stMetricDelta"] { font-size:.72rem; }

        /* Horizontal nav: text-first, no pill backgrounds. */
        div[role="radiogroup"] {
            gap:22px !important;
            border-bottom:1px solid var(--border);
            margin-bottom:18px;
        }
        div[role="radiogroup"] > label {
            padding:0 0 10px !important;
            margin:0 !important;
            border-radius:0 !important;
            color:var(--muted) !important;
            font-weight:650 !important;
        }
        div[role="radiogroup"] > label:has(input:checked) {
            color:var(--text) !important;
            border-bottom:2px solid var(--text);
        }
        div[role="radiogroup"] [data-testid="stMarkdownContainer"] p { font-size:.84rem; }
        div[role="radiogroup"] > label > div:first-child { display:none; }

        [data-baseweb="select"] > div,
        .stTextInput input,
        .stDateInput input {
            border-color:var(--border-strong) !important;
            border-radius:8px !important;
            box-shadow:none !important;
        }
        [data-testid="stFileUploaderDropzone"] {
            background:var(--surface-soft);
            border:1px dashed var(--border-strong);
            border-radius:10px;
        }

        div[data-testid="stDataFrame"] {
            border:1px solid var(--border);
            border-radius:9px;
            overflow:hidden;
            background:#fff;
        }

        .stButton > button,
        .stDownloadButton > button {
            border-radius:8px;
            min-height:40px;
            font-weight:680;
            box-shadow:none;
        }
        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"] {
            background:#111827;
            border-color:#111827;
            color:#fff;
        }
        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button[kind="primary"]:hover {
            background:#1F2937;
            border-color:#1F2937;
        }

        .empty-state {
            border:1px dashed var(--border-strong);
            border-radius:12px;
            padding:36px 24px;
            text-align:center;
            color:var(--muted);
            background:#fff;
        }
        .empty-state strong { display:block; color:var(--text); font-size:1rem; margin-bottom:5px; }

        @media (max-width: 900px) {
            .block-container { padding-left: 1rem; padding-right: 1rem; }
            .topbar { align-items:flex-start; }
            div[role="radiogroup"] { gap:13px !important; overflow-x:auto; flex-wrap:nowrap !important; }
            div[role="radiogroup"] > label { white-space:nowrap; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _money(valor):
    try:
        return f"${float(valor):,.2f}"
    except Exception:
        return "$0.00"


def _topbar():
    st.markdown(
        f"""
        <div class="topbar">
          <div class="brand">
            <span class="brand-mark"></span>
            <div>
              <div class="brand-title">Auditoría ARPON</div>
              <div class="brand-sub">Hotel Quartz · Clientes y Proveedores</div>
            </div>
          </div>
          <div class="version">{APP_VERSION}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _page_head(title, subtitle=""):
    st.markdown(
        f'<div class="page-head"><h1>{title}</h1><p>{subtitle}</p></div>',
        unsafe_allow_html=True,
    )


def _section(title, subtitle=""):
    st.markdown(
        f'<div class="section-head"><h2>{title}</h2><p>{subtitle}</p></div>',
        unsafe_allow_html=True,
    )


def _status(kind, title, detail=""):
    texto = f"<strong>{title}</strong>"
    if detail:
        texto += f" &nbsp;·&nbsp; {detail}"
    st.markdown(
        f'<div class="status-line {kind}"><span class="status-dot"></span><div>{texto}</div></div>',
        unsafe_allow_html=True,
    )


def _chips(items):
    html = ''.join(f'<span class="meta-chip">{x}</span>' for x in items if x)
    st.markdown(f'<div class="meta-row">{html}</div>', unsafe_allow_html=True)


def _empty(title, detail):
    st.markdown(
        f'<div class="empty-state"><strong>{title}</strong>{detail}</div>',
        unsafe_allow_html=True,
    )


def _aplicar_tipo(df, tipos):
    if df is None:
        return pd.DataFrame()
    if "tipo_cuenta" in df.columns and tipos is not None:
        return df[df["tipo_cuenta"].isin(tipos)].copy()
    return df


def _tabla(df, cols=None, height=520):
    if cols is not None:
        cols = [c for c in cols if c in df.columns]
        df = df[cols]
    st.dataframe(df, use_container_width=True, hide_index=True, height=height)


def _render_landing(uploaded_files):
    _page_head(
        "Audita auxiliares ARPON sin convertirlos en otro sistema",
        "Carga Clientes, Proveedores o ambos. El motor valida primero la fuente y solo después genera conciliaciones, aging y hallazgos.",
    )
    _chips(["Clientes / CxC", "Proveedores / CxP", "XLSX · XLSM · XLS · CSV"])
    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("#### Archivos de trabajo")
        st.caption("Puedes cargar uno o varios periodos. Si existe traslape de fechas en la misma cuenta, el proceso se detiene.")
        # El uploader real se renderiza en main; aquí solo se mantiene el espacio de contexto.
        if not uploaded_files:
            st.info("Selecciona los auxiliares en el cargador superior para iniciar la validación.", icon=None)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.metric("Validación", "Movimiento a movimiento")
    c2.metric("Conciliación", "Documento + evidencia")
    c3.metric("Salida", "Excel auditable")


def main():
    st.set_page_config(
        page_title="Auditoría ARPON · Hotel Quartz",
        layout="wide",
        page_icon="Q",
        initial_sidebar_state="collapsed",
    )
    _inyectar_estilos_ui()
    _topbar()

    # Carga siempre visible, pero contenida y compacta.
    with st.expander("Archivos ARPON", expanded=True):
        uploaded_files = st.file_uploader(
            "Selecciona auxiliares",
            type=["xlsx", "xls", "xlsm", "csv"],
            accept_multiple_files=True,
            help="Carga Clientes, Proveedores o ambos. Los periodos de una misma cuenta no deben traslaparse.",
            label_visibility="collapsed",
            key="archivos_arpon",
        )
        if uploaded_files:
            st.caption(" · ".join(f.name for f in uploaded_files))

    if not uploaded_files:
        _render_landing(uploaded_files)
        return

    movs_lista, resumen_lista, diags, errores = [], [], [], []
    with st.spinner("Validando auxiliares…"):
        for uf in uploaded_files:
            try:
                movs_i, resumen_i, diag_i = procesar_archivo_engine(uf.getvalue(), uf.name)
                movs_lista.append(movs_i)
                resumen_lista.append(resumen_i)
                diags.append(diag_i)
            except Exception as e:
                errores.append(f"{uf.name}: {e}")

    if errores:
        _page_head("La carga no pasó la validación", "No se muestran resultados parciales cuando un archivo falla.")
        for error in errores:
            _status("danger", "Archivo rechazado", error)
        return

    movs = pd.concat(movs_lista, ignore_index=True)
    resumen = pd.concat(resumen_lista, ignore_index=True)
    diag_df = pd.DataFrame(diags)

    solapamientos = detectar_solapamientos_periodos(resumen)
    if not solapamientos.empty:
        _page_head("Hay periodos superpuestos", "La misma cuenta contiene días repetidos entre archivos; continuar podría duplicar movimientos.")
        _status("danger", "Proceso detenido", "Carga periodos consecutivos sin traslape para la misma cuenta.")
        _tabla(solapamientos, height=360)
        return

    resumen_nat = detectar_naturaleza(resumen, movs)
    movs = aplicar_naturaleza_a_movimientos(movs, resumen_nat)
    movs = marcar_duplicados_exactos(movs)
    df_cruces_ref = detectar_cruces_por_referencia(movs)
    df_evidencia = detectar_coincidencias_por_evidencia(movs)
    movs = marcar_movimientos_conciliacion(movs, df_cruces_ref, df_evidencia)
    df_audit = analizar_saldos(movs, resumen_nat)

    fmax = movs["fecha"].max()
    corte_default = fmax.date() if pd.notna(fmax) else pd.Timestamp.now().date()

    n_archivos = len(diag_df)
    n_cuentas = len(df_audit)
    n_movs = len(movs)
    n_reparaciones = int(movs["fila_reparada"].sum())
    saldo_clientes = float(df_audit.loc[df_audit["tipo_cuenta"].eq("CLIENTES"), "saldo_final_pendiente"].sum())
    saldo_proveedores = float(df_audit.loc[df_audit["tipo_cuenta"].eq("PROVEEDORES"), "saldo_final_pendiente"].sum())
    n_partidas_conciliadas = int(movs["conciliacion_estado"].eq("CONCILIADO").sum())
    n_partidas_revisar = int(movs["conciliacion_estado"].eq("REVISAR").sum())
    n_documentos_conciliados = int(
        movs[movs["conciliacion_estado"].eq("CONCILIADO") & movs["es_documento_conciliable"]]
        [["cuenta_logica_uid", "referencia_norm"]].drop_duplicates().shape[0]
    )

    # Filtros: fuera del camino principal.
    tipos_disponibles = sorted(movs["tipo_cuenta"].dropna().astype(str).unique())
    opciones_empresas = sorted(movs["empresa"].fillna("").astype(str).loc[lambda s: s.str.strip().ne("")].unique())
    opciones_cuentas = sorted(movs["meta_codigo"].fillna("").astype(str).unique())
    opciones_archivos = sorted(movs["archivo"].fillna("").astype(str).unique())
    fecha_min, fecha_max = movs["fecha"].min(), movs["fecha"].max()

    with st.expander("Filtros y fecha de corte", expanded=False):
        f1, f2, f3, f4 = st.columns(4)
        tipos_filtro = f1.multiselect("Tipo de cuenta", tipos_disponibles, default=tipos_disponibles)
        empresas_filtro = f2.multiselect("Empresa", opciones_empresas, default=opciones_empresas) if opciones_empresas else None
        cuentas_filtro = f3.multiselect("Cuenta", opciones_cuentas, default=opciones_cuentas)
        archivos_filtro = f4.multiselect("Archivo", opciones_archivos, default=opciones_archivos)

        g1, g2, g3 = st.columns([1.15, 1.55, 1])
        if pd.notna(fecha_min) and pd.notna(fecha_max):
            rango_fechas = g1.date_input(
                "Fecha de movimiento",
                value=(fecha_min.date(), fecha_max.date()),
                min_value=fecha_min.date(), max_value=fecha_max.date(),
            )
            if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
                fecha_desde_filtro, fecha_hasta_filtro = rango_fechas
            else:
                fecha_desde_filtro = fecha_hasta_filtro = rango_fechas
        else:
            fecha_desde_filtro = fecha_hasta_filtro = None
        busqueda_filtro = g2.text_input("Buscar", placeholder="Documento, póliza, concepto, cuenta…")
        corte = g3.date_input("Corte de antigüedad", value=corte_default)

    folios = analizar_folios(movs, corte)
    filtros_tabla = {
        "empresas": empresas_filtro if opciones_empresas else None,
        "archivos": archivos_filtro,
        "cuentas": cuentas_filtro,
        "fecha_desde": fecha_desde_filtro,
        "fecha_hasta": fecha_hasta_filtro,
        "busqueda": busqueda_filtro,
    }
    movs_vista = _aplicar_tipo(aplicar_filtros_tabla(movs, **filtros_tabla), tipos_filtro)
    audit_vista = _aplicar_tipo(aplicar_filtros_tabla(df_audit, **filtros_tabla), tipos_filtro)
    folios_vista = _aplicar_tipo(aplicar_filtros_tabla(folios, **filtros_tabla), tipos_filtro)
    cruces_ref_vista = _aplicar_tipo(aplicar_filtros_tabla(df_cruces_ref, **filtros_tabla), tipos_filtro)
    evidencia_vista = _aplicar_tipo(aplicar_filtros_tabla(df_evidencia, **filtros_tabla), tipos_filtro)
    diag_vista = aplicar_filtros_tabla(diag_df, **filtros_tabla)

    sin_ref_movs = movs_vista[~movs_vista["tiene_referencia"]].copy()
    refs_rec = movs_vista[movs_vista["referencia_recuperada"]].copy()
    negativos = movs_vista[(movs_vista["cargos"] < 0) | (movs_vista["abonos"] < 0)].copy()
    duplicados = movs_vista[movs_vista["posible_duplicado_exacto"]].copy()
    reparadas = movs_vista[movs_vista["fila_reparada"]].copy()
    descuadres = audit_vista[audit_vista["descuadre_origen"].abs() > UMBRAL_TOLERANCIA].copy()
    contrarios = folios_vista[folios_vista["tipo_saldo"].str.contains("contrario", case=False, na=False)].copy()
    viejos = folios_vista[folios_vista["antiguedad_observada"].eq("90+")].copy()
    n_hallazgos_accion = len(descuadres) + len(sin_ref_movs) + len(duplicados) + len(contrarios)

    _chips([
        f"{n_archivos} archivo(s)",
        f"{n_cuentas} cuenta(s)",
        f"{n_movs:,} movimientos",
        f"{n_reparaciones} reconstrucción(es)",
    ])
    st.markdown("<div style='height:11px'></div>", unsafe_allow_html=True)

    paginas = ["Resumen", "Hallazgos", "Conciliación", "Documentos", "Movimientos", "Diagnóstico", "Exportar"]
    pagina = st.radio("Sección", paginas, horizontal=True, label_visibility="collapsed", key="nav_principal")

    if pagina == "Resumen":
        _page_head("Resumen", "Estado general de la carga y saldos pendientes normalizados.")
        amarres_false = int(diag_df["amarre_gran_total"].eq(False).sum())
        if amarres_false == 0:
            _status("success", "Carga validada", "Estructura, secuencia y totales disponibles superaron las validaciones.")
        else:
            _status("warning", "Carga con observaciones", f"{amarres_false} archivo(s) sin amarre independiente de gran total.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Clientes · pendiente de cobro", _money(saldo_clientes))
        c2.metric("Proveedores · pendiente de pago", _money(saldo_proveedores))
        c3.metric("Documentos conciliados", f"{n_documentos_conciliados:,}", f"{n_partidas_conciliadas:,} partidas")
        c4.metric("Requieren atención", f"{n_hallazgos_accion:,}", f"{n_partidas_revisar:,} conciliaciones con remanente")

        left, right = st.columns([1.38, .62])
        with left:
            _section("Antigüedad de documentos abiertos", "Saldo pendiente por primera fecha observada del documento.")
            orden = ["0-30", "31-60", "61-90", "90+"]
            positivos = folios_vista[folios_vista["saldo_natural"] > 0].copy()
            aging = (
                positivos.groupby("antiguedad_observada")["saldo_natural"]
                .agg(num_documentos="count", saldo="sum")
                .reindex(orden).fillna(0).reset_index()
            )
            fig = go.Figure(go.Bar(
                x=aging["antiguedad_observada"], y=aging["saldo"],
                text=[_money(x) for x in aging["saldo"]], textposition="outside",
                marker_color="#344054",
            ))
            fig.update_layout(
                height=315, margin=dict(l=4, r=4, t=14, b=4), showlegend=False,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(title=None, showgrid=False, zeroline=False),
                yaxis=dict(title=None, gridcolor="#F2F4F7", tickprefix="$", separatethousands=True),
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        with right:
            _section("Atención", "Solo elementos con posible acción.")
            if len(descuadres): _status("danger", f"{len(descuadres)} cuenta(s) con descuadre")
            if len(sin_ref_movs): _status("warning", f"{len(sin_ref_movs)} movimiento(s) sin referencia")
            if len(duplicados): _status("warning", f"{len(duplicados)} posible(s) duplicado(s)")
            if len(contrarios): _status("warning", f"{len(contrarios)} documento(s) con saldo contrario")
            if not n_hallazgos_accion: _status("success", "Sin hallazgos críticos con los filtros actuales")
            if n_reparaciones: _status("info", f"{n_reparaciones} partida(s) ARPON reconstruida(s)", "Disponibles en Diagnóstico.")

        _section("Cuentas", "Vista compacta del estado por cuenta contable.")
        cols = ["tipo_cuenta", "empresa", "meta_codigo", "meta_nombre", "estado", "saldo_inicial_pendiente", "total_cargos", "total_abonos", "saldo_final_pendiente", "n_sin_referencia", "descuadre_origen"]
        _tabla(audit_vista, cols, height=370)

    elif pagina == "Hallazgos":
        _page_head("Hallazgos", "Revisa una categoría a la vez; la herramienta no mezcla observaciones de distinta naturaleza.")
        categorias = {
            "Descuadres": descuadres,
            "Sin referencia": sin_ref_movs,
            "Duplicados": duplicados,
            "Montos negativos / reversos": negativos,
            "Referencias recuperadas": refs_rec,
            "Saldos contrarios": contrarios,
            "Documentos 90+": viejos,
            "Reconstrucciones ARPON": reparadas,
        }
        h1, h2 = st.columns([.72, .28])
        categoria = h1.selectbox("Categoría", list(categorias.keys()))
        actual = categorias[categoria]
        h2.metric("Resultados", f"{len(actual):,}")
        if actual.empty:
            _empty("Sin resultados", "No hay elementos en esta categoría con los filtros actuales.")
        else:
            _tabla(actual, height=620)

    elif pagina == "Conciliación":
        _page_head("Conciliación", "Partidas relacionadas por documento o evidencia determinística.")
        c1, c2, c3 = st.columns(3)
        c1.metric("Partidas conciliadas", f"{n_partidas_conciliadas:,}")
        c2.metric("Partidas a revisar", f"{n_partidas_revisar:,}")
        c3.metric("Documentos conciliados", f"{n_documentos_conciliados:,}")

        partidas = movs_vista[movs_vista["conciliacion_marcada"]].copy()
        if partidas.empty:
            _empty("Sin partidas marcadas", "No se detectaron relaciones de conciliación con los filtros actuales.")
        else:
            f1, f2 = st.columns(2)
            estados = sorted(partidas["conciliacion_estado"].dropna().unique())
            niveles = sorted(partidas["conciliacion_nivel"].dropna().loc[lambda s: s.astype(str).str.strip().ne("")].unique())
            estados_sel = f1.multiselect("Estado", estados, default=estados)
            niveles_sel = f2.multiselect("Nivel de evidencia", niveles, default=niveles)
            p = partidas[partidas["conciliacion_estado"].isin(estados_sel) & partidas["conciliacion_nivel"].isin(niveles_sel)].copy()
            cols = ["conciliacion_estado", "conciliacion_nivel", "conciliacion_criterio", "conciliacion_codigo", "tipo_cuenta", "archivo", "fila_origen", "fecha", "meta_codigo", "poliza", "referencia_original", "concepto", "cargos", "abonos", "efecto_natural"]
            _tabla(p, cols, height=610)

        with st.expander("Cruces por mismo documento"):
            if cruces_ref_vista.empty: st.caption("Sin cruces para los filtros actuales.")
            else: _tabla(cruces_ref_vista, height=380)
        with st.expander("Coincidencias por evidencia"):
            if evidencia_vista.empty: st.caption("Sin coincidencias por evidencia para los filtros actuales.")
            else: _tabla(evidencia_vista, height=380)

    elif pagina == "Documentos":
        _page_head("Documentos abiertos", "Folios numéricos, prefijados y documentos alfanuméricos con saldo vivo.")
        orden = ["0-30", "31-60", "61-90", "90+"]
        positivos = folios_vista[folios_vista["saldo_natural"] > 0].copy()
        aging = positivos.groupby("antiguedad_observada")["saldo_natural"].agg(num_documentos="count", saldo="sum").reindex(orden).fillna(0).reset_index()
        cols_metric = st.columns(4)
        for col, bucket in zip(cols_metric, orden):
            row = aging[aging["antiguedad_observada"].eq(bucket)].iloc[0]
            col.metric(f"{bucket} días", _money(row["saldo"]), f"{int(row['num_documentos']):,} documentos")

        if folios_vista.empty:
            _empty("Sin documentos abiertos", "No hay saldos documentales vivos con los filtros actuales.")
        else:
            d1, d2 = st.columns(2)
            nat_opciones = sorted(folios_vista["naturaleza"].dropna().unique())
            nat_sel = d1.multiselect("Naturaleza", nat_opciones, default=nat_opciones)
            edades_sel = d2.multiselect("Antigüedad", orden, default=orden)
            fv = folios_vista[folios_vista["naturaleza"].isin(nat_sel) & folios_vista["antiguedad_observada"].isin(edades_sel)].copy()
            cols = ["tipo_cuenta", "empresa", "meta_codigo", "meta_nombre", "referencia_norm", "primera_fecha", "ultima_fecha", "dias", "antiguedad_observada", "cargos", "abonos", "saldo_natural", "tipo_saldo"]
            _tabla(fv, cols, height=650)

    elif pagina == "Movimientos":
        _page_head("Movimientos", "Detalle completo del auxiliar después de la lectura certificada.")
        c1, c2, c3 = st.columns(3)
        c1.metric("Visibles", f"{len(movs_vista):,}")
        c2.metric("Con referencia", f"{int(movs_vista['tiene_referencia'].sum()):,}")
        c3.metric("Sin referencia", f"{int((~movs_vista['tiene_referencia']).sum()):,}")
        cols = ["archivo", "fila_origen", "fecha", "tipo_cuenta", "meta_codigo", "poliza", "referencia_original", "referencia_norm", "referencia_tipo", "concepto", "cargos", "abonos", "saldo_acumulado", "efecto_natural", "conciliacion_estado"]
        _tabla(movs_vista, cols, height=690)

    elif pagina == "Diagnóstico":
        _page_head("Diagnóstico", "Trazabilidad técnica de la lectura, las ecuaciones ARPON y las reconstrucciones.")
        with st.expander("Archivos y validación de origen", expanded=True):
            _tabla(diag_vista, height=350)
        with st.expander("Naturaleza y ecuación ARPON"):
            diag_cols = ["sistema_origen", "empresa", "tipo_cuenta", "archivo", "meta_codigo", "meta_nombre", "naturaleza", "naturaleza_confianza", "saldo_inicial", "total_cargos", "total_abonos", "esperado_arpon", "saldo_final_aux", "error_arpon", "saldo_final_pendiente", "ecuacion_saldo_fuente", "n_errores_saldo_secuencia", "max_error_saldo_secuencia", "error_ultimo_saldo_vs_total", "n_filas_reconstruidas"]
            _tabla(audit_vista, diag_cols, height=450)
        with st.expander("Reconstrucciones ARPON", expanded=bool(len(reparadas))):
            if reparadas.empty:
                st.caption("No hubo partidas fragmentadas que reconstruir.")
            else:
                cols = ["archivo", "fila_origen", "fila_continuacion", "fecha", "meta_codigo", "poliza", "referencia_original", "concepto", "cargos", "abonos", "saldo_acumulado"]
                _tabla(reparadas, cols, height=360)

    elif pagina == "Exportar":
        _page_head("Exportar", "Genera el expediente reproducible y los auxiliares ARPON marcados.")
        ejecucion_rows = []
        fecha_ejecucion = pd.Timestamp.now(tz="UTC").isoformat()
        for uf in uploaded_files:
            data_bytes = uf.getvalue()
            ejecucion_rows.append({
                "version_motor": APP_VERSION,
                "fecha_ejecucion_utc": fecha_ejecucion,
                "archivo": uf.name,
                "sha256": hashlib.sha256(data_bytes).hexdigest(),
                "bytes": len(data_bytes),
                "tolerancia_contable": UMBRAL_TOLERANCIA,
                "umbral_documento": UMBRAL_FOLIO,
            })
        ejecucion_df = pd.DataFrame(ejecucion_rows)
        reparaciones_df = movs[movs["fila_reparada"]].copy()
        export_tables = {
            "Ejecucion": ejecucion_df,
            "Semaforo": df_audit,
            "Documentos": folios,
            "Movimientos": movs,
            "Reparaciones_ARPON": reparaciones_df,
            "Cruces_documento": df_cruces_ref,
            "Cruces_evidencia": df_evidencia,
            "Diagnostico": diag_df,
        }
        e1, e2 = st.columns(2)
        with e1:
            with st.container(border=True):
                st.markdown("#### Auditoría completa")
                st.caption("Ejecución, saldos, documentos, movimientos, reparaciones, cruces y diagnóstico.")
                st.download_button(
                    "Descargar Excel maestro",
                    data=to_excel_workbook(export_tables),
                    file_name="auditoria_master_saldos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary", use_container_width=True,
                )
        with e2:
            with st.container(border=True):
                st.markdown("#### Auxiliares marcados")
                st.caption("Conservan el origen y agregan marcas de conciliación, revisión y reconstrucción.")
                if n_partidas_conciliadas or n_partidas_revisar or n_reparaciones:
                    archivos_origen = [(uf.name, uf.getvalue()) for uf in uploaded_files]
                    data_marcada, nombre_marcado, mime_marcado = construir_descarga_auxiliares_marcados(archivos_origen, movs)
                    st.download_button(
                        "Descargar auxiliares",
                        data=data_marcada, file_name=nombre_marcado, mime=mime_marcado,
                        type="primary", use_container_width=True,
                    )
                else:
                    st.button("Sin marcas para exportar", disabled=True, use_container_width=True)
        _status("info", "Expediente reproducible", "La hoja Ejecución conserva versión, fecha UTC, SHA-256 y tolerancias del motor.")


if __name__ == "__main__":
    main()
