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
APP_VERSION = "5.0 ARPON · HOTEL QUARTZ · CLIENTES + PROVEEDORES"
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
# 5. UI
# ==============================================================================

def main():
    st.set_page_config(
        page_title="Auditoría ARPON · Clientes y Proveedores · Hotel Quartz",
        layout="wide",
        page_icon="🛡️",
    )

    st.title("🛡️ Auditoría ARPON · Clientes y Proveedores · Hotel Quartz")
    st.caption(f"Motor v{APP_VERSION}")

    st.markdown(
        """
        Motor exclusivo para auxiliares **ARPON de Clientes y Proveedores de Hotel Quartz**.

        - valida la estructura **Póliza | Fecha | Docto. | Concepto | Cargo | Abono | Saldo**;
        - identifica empresa, periodo, cuenta y saldo inicial desde el propio reporte;
        - valida **cargos, abonos y saldo acumulado movimiento por movimiento**;
        - identifica **Clientes / Proveedores** y normaliza el saldo pendiente sin alterar la ecuación ARPON;
        - conserva y normaliza **folios, documentos numéricos y documentos alfanuméricos** sin destruir identificadores;
        - puede recuperar un folio desde **Concepto** cuando Docto. está vacío;
        - identifica movimientos sin referencia, reversos y posibles duplicados;
        - los cruces se realizan únicamente dentro de la **misma empresa ARPON**;
        - reconstruye de forma controlada partidas ARPON fragmentadas y deja trazabilidad;
        - genera una copia del auxiliar con **color y código de conciliación/auditoría**.
        """
    )

    uploaded_files = st.file_uploader(
        "📂 Sube auxiliares ARPON de Clientes y/o Proveedores (Excel o CSV)",
        type=["xlsx", "xls", "xlsm", "csv"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Esperando archivo(s)...")
        return

    movs_lista = []
    resumen_lista = []
    diags = []
    errores = []

    with st.spinner("Procesando y validando auxiliares..."):
        for uf in uploaded_files:
            try:
                movs_i, resumen_i, diag_i = procesar_archivo_engine(
                    uf.getvalue(), uf.name
                )
                movs_lista.append(movs_i)
                resumen_lista.append(resumen_i)
                diags.append(diag_i)
            except Exception as e:
                errores.append(f"**{uf.name}:** {e}")

    if errores:
        st.error(
            "⚠️ **No se certifica la lectura. Corrige o revisa estos archivos antes "
            "de usar resultados:**\n\n"
            + "\n\n".join(f"- {x}" for x in errores)
        )
        st.stop()

    movs = pd.concat(movs_lista, ignore_index=True)
    resumen = pd.concat(resumen_lista, ignore_index=True)

    solapamientos = detectar_solapamientos_periodos(resumen)
    if not solapamientos.empty:
        st.error(
            "⛔ Se detectaron periodos superpuestos para la misma cuenta. "
            "Cargar periodos traslapados duplicaría movimientos y puede crear conciliaciones falsas. "
            "Carga archivos sin días repetidos para esa cuenta."
        )
        st.dataframe(solapamientos, use_container_width=True, hide_index=True)
        st.stop()

    # Evitar que cargar dos veces la misma cuenta pase desapercibido.
    repetidas = (
        resumen.groupby(["empresa_uid", "meta_codigo"])["archivo"]
        .nunique()
        .loc[lambda s: s > 1]
    )
    if not repetidas.empty:
        st.warning(
            "ℹ️ Hay cuentas presentes en más de un archivo. Los periodos no se solapan, "
            "por lo que el motor los analizará como continuidad histórica: "
            + ", ".join(str(x) for x in repetidas.index)
        )

    resumen_nat = detectar_naturaleza(resumen, movs)
    movs = aplicar_naturaleza_a_movimientos(movs, resumen_nat)
    movs = marcar_duplicados_exactos(movs)
    df_cruces_ref = detectar_cruces_por_referencia(movs)
    df_evidencia = detectar_coincidencias_por_evidencia(movs)
    movs = marcar_movimientos_conciliacion(
        movs, df_cruces_ref, df_evidencia
    )
    df_audit = analizar_saldos(movs, resumen_nat)

    # --------------------------------------------------------------------------
    # Validación visible de lectura
    # --------------------------------------------------------------------------
    st.divider()
    st.subheader("✅ Validación de lectura")

    diag_df = pd.DataFrame(diags)

    st.caption("Sistema contable: ARPON")
    empresas_arpon = sorted(
        diag_df["empresa"].dropna().astype(str)
        .loc[lambda x: x.str.strip().ne("")].unique()
    )
    if empresas_arpon:
        st.caption("Empresa detectada: " + " | ".join(empresas_arpon))

    n_archivos = len(diag_df)
    n_cuentas = len(df_audit)
    n_movs = len(movs)

    amarres_false = diag_df["amarre_gran_total"].eq(False).sum()
    if amarres_false:
        st.warning(
            f"{amarres_false} archivo(s) no amarran la suma de saldos por cuenta "
            "contra el saldo final reportado por ARPON. Revisa si el reporte contiene "
            "agrupaciones adicionales."
        )
    else:
        st.success(
            f"Lectura estructural validada: **{n_archivos} archivo(s)** · "
            f"**{n_cuentas} cuenta(s)** · **{n_movs:,} movimientos** · "
            f"**{int(diag_df['n_filas_reconstruidas'].fillna(0).sum()):,} fila(s) reconstruida(s)**. "
            "La estructura, los totales ARPON disponibles y las secuencias de saldo fueron validados."
        )

    for _, d in diag_df.iterrows():
        gt_txt = (
            f"${d['gran_total']:,.2f}"
            if pd.notna(d.get("gran_total"))
            else "no detectado"
        )
        amarre = d.get("amarre_gran_total")
        if pd.isna(amarre):
            estado_amarre = "ℹ️"
        elif bool(amarre):
            estado_amarre = "✅"
        else:
            estado_amarre = "⚠️"
        st.caption(
            f"{estado_amarre} {d['archivo']} · {'ARPON'} [{d.get('formato', 'N/D')}]: "
            f"{int(d['n_headers'])} cuenta(s), {int(d['n_movs']):,} movimientos, "
            f"Total {d.get('origen_gran_total', 'N/D')} {gt_txt} · reconstruidas {int(d.get('n_filas_reconstruidas', 0))}."
        )

    # --------------------------------------------------------------------------
    # KPIs
    # --------------------------------------------------------------------------
    saldo_total = float(df_audit["saldo_final_pendiente"].sum())
    saldo_clientes = float(
        df_audit.loc[df_audit["tipo_cuenta"].eq("CLIENTES"), "saldo_final_pendiente"].sum()
    )
    saldo_proveedores = float(
        df_audit.loc[df_audit["tipo_cuenta"].eq("PROVEEDORES"), "saldo_final_pendiente"].sum()
    )
    bruto_sin_ref = df_audit["importe_bruto_sin_referencia"].sum()
    descuadre_abs = df_audit["descuadre_origen"].abs().sum()
    n_sin_ref = int(df_audit["n_sin_referencia"].sum())
    n_revisar = int((df_audit["estado"] != "🟢 OK").sum())

    n_folios_conciliados = int(
        movs[
            movs["conciliacion_estado"].eq("CONCILIADO")
            & movs["es_documento_conciliable"]
        ][["cuenta_logica_uid", "referencia_norm"]].drop_duplicates().shape[0]
    )
    n_evidencias = (
        int(df_evidencia["evidencia_id"].nunique())
        if not df_evidencia.empty else 0
    )
    n_partidas_conciliadas = int(
        movs["conciliacion_estado"].eq("CONCILIADO").sum()
    )
    n_partidas_revisar = int(
        movs["conciliacion_estado"].eq("REVISAR").sum()
    )

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Saldo pendiente normalizado", f"${saldo_total:,.2f}")
    k2.metric(
        "Movs sin referencia",
        f"{n_sin_ref:,}",
        help=f"Importe bruto involucrado: ${bruto_sin_ref:,.2f}",
    )
    k3.metric(
        "Descuadre absoluto",
        f"${descuadre_abs:,.2f}",
        help="Suma de valores absolutos por cuenta; evita compensar + y -.",
    )
    k4.metric("Documentos conciliados", n_folios_conciliados)
    k5.metric("Cruces por evidencia", n_evidencias)
    k6.metric("Cuentas a revisar", n_revisar)
    n_reparaciones = int(movs["fila_reparada"].sum())
    st.caption(
        f"Pendiente normalizado · Clientes: ${saldo_clientes:,.2f} · "
        f"Proveedores: ${saldo_proveedores:,.2f}. Marcas: "
        f"{n_partidas_conciliadas:,} conciliada(s), {n_partidas_revisar:,} a revisar y "
        f"{n_reparaciones:,} fila(s) ARPON reconstruida(s)."
    )

    # Fecha de corte
    fmax = movs["fecha"].max()
    corte_default = (
        fmax.date() if pd.notna(fmax) else pd.Timestamp.now().date()
    )
    corte = st.date_input(
        "📅 Fecha de corte para antigüedad observada",
        value=corte_default,
        help=(
            "La antigüedad se calcula desde la primera fecha observada del folio. "
            "No equivale a días vencidos si no existe fecha de vencimiento."
        ),
    )
    folios = analizar_folios(movs, corte)

    # --------------------------------------------------------------------------
    # Filtros generales de tablas
    # --------------------------------------------------------------------------
    st.divider()
    with st.expander("🎛️ Filtros de las tablas", expanded=True):
        st.caption(
            "Estos filtros se aplican a las tablas en todas las pestañas. Los "
            "indicadores superiores y los archivos exportados conservan los "
            "resultados completos."
        )
        fg1, fg2, fg3 = st.columns(3)
        opciones_empresas = sorted(
            movs["empresa"].fillna("").astype(str)
            .loc[lambda s: s.str.strip().ne("")].unique()
        )
        opciones_archivos = sorted(
            movs["archivo"].fillna("").astype(str).unique()
        )
        opciones_cuentas = sorted(
            movs["meta_codigo"].fillna("").astype(str).unique()
        )
        empresas_filtro = fg1.multiselect(
            "Empresa",
            opciones_empresas,
            default=opciones_empresas,
            key="filtro_global_empresa",
        )
        archivos_filtro = fg2.multiselect(
            "Archivo",
            opciones_archivos,
            default=opciones_archivos,
            key="filtro_global_archivo",
        )
        cuentas_filtro = fg3.multiselect(
            "Cuenta contable",
            opciones_cuentas,
            default=opciones_cuentas,
            key="filtro_global_cuenta",
        )

        fg4, fg5 = st.columns([1, 2])
        fecha_min = movs["fecha"].min()
        fecha_max = movs["fecha"].max()
        if pd.notna(fecha_min) and pd.notna(fecha_max):
            rango_fechas = fg4.date_input(
                "Fecha del movimiento",
                value=(fecha_min.date(), fecha_max.date()),
                min_value=fecha_min.date(),
                max_value=fecha_max.date(),
                key="filtro_global_fecha",
            )
            if isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
                fecha_desde_filtro, fecha_hasta_filtro = rango_fechas
            else:
                fecha_desde_filtro = rango_fechas
                fecha_hasta_filtro = rango_fechas
        else:
            fecha_desde_filtro = None
            fecha_hasta_filtro = None

        busqueda_filtro = fg5.text_input(
            "Buscar en las tablas",
            placeholder=(
                "Folio, póliza, concepto, cuenta, archivo, estado o código..."
            ),
            key="filtro_global_busqueda",
        )

    filtros_tabla = {
        "empresas": empresas_filtro if opciones_empresas else None,
        "archivos": archivos_filtro,
        "cuentas": cuentas_filtro,
        "fecha_desde": fecha_desde_filtro,
        "fecha_hasta": fecha_hasta_filtro,
        "busqueda": busqueda_filtro,
    }
    movs_vista = aplicar_filtros_tabla(movs, **filtros_tabla)
    audit_vista = aplicar_filtros_tabla(df_audit, **filtros_tabla)
    folios_vista = aplicar_filtros_tabla(folios, **filtros_tabla)
    cruces_ref_vista = aplicar_filtros_tabla(
        df_cruces_ref, **filtros_tabla
    )
    evidencia_vista = aplicar_filtros_tabla(
        df_evidencia, **filtros_tabla
    )
    diag_vista = aplicar_filtros_tabla(diag_df, **filtros_tabla)
    st.caption(
        f"Resultado de filtros: {len(movs_vista):,} de {len(movs):,} "
        f"movimiento(s) · {len(audit_vista):,} de {len(df_audit):,} cuenta(s)."
    )

    # --------------------------------------------------------------------------
    # Pestañas
    # --------------------------------------------------------------------------
    tabs = st.tabs(
        [
            "🔎 Hallazgos",
            "🚦 Semáforo",
            "📑 Documentos",
            "✅ Conciliación marcada",
            "🏷️ Referencias",
            "📉 Gráficos",
            "🧪 Diagnóstico",
        ]
    )

    # --------------------------------------------------------------------------
    # Hallazgos
    # --------------------------------------------------------------------------
    with tabs[0]:
        st.subheader("🔎 Hallazgos priorizados")
        st.caption(
            "Los hallazgos son independientes. Una cuenta puede tener más de uno."
        )

        sin_ref_movs = movs_vista[~movs_vista["tiene_referencia"]].copy()
        refs_rec = movs_vista[movs_vista["referencia_recuperada"]].copy()
        negativos = movs_vista[
            (movs_vista["cargos"] < 0) | (movs_vista["abonos"] < 0)
        ].copy()
        duplicados = movs_vista[
            movs_vista["posible_duplicado_exacto"]
        ].copy()
        descuadres = audit_vista[
            audit_vista["descuadre_origen"].abs() > UMBRAL_TOLERANCIA
        ].copy()
        contrarios = folios_vista[
            folios_vista["tipo_saldo"].str.contains(
                "contrario", case=False, na=False
            )
        ].copy()
        viejos = folios_vista[
            folios_vista["antiguedad_observada"].eq("90+")
        ].copy()

        h1, h2, h3, h4, h5, h6 = st.columns(6)
        h1.metric("Cuentas descuadre", len(descuadres))
        h2.metric("Movs sin ref", len(sin_ref_movs))
        h3.metric("Refs recuperadas", len(refs_rec))
        h4.metric("Montos negativos", len(negativos))
        h5.metric("Posibles duplicados", len(duplicados))
        h6.metric("Documentos 90+ observados", len(viejos))

        if len(descuadres):
            st.markdown("#### 🟠 Descuadre contra el saldo final reportado por ARPON")
            st.dataframe(
                descuadres[
                    [
                        "sistema_origen", "empresa", "archivo", "meta_codigo",
                        "meta_nombre", "naturaleza", "saldo_final_aux", "saldo_esperado_motor",
                        "descuadre_origen"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(sin_ref_movs):
            st.markdown("#### 🔴 Movimientos realmente sin referencia")
            st.caption(
                "No había referencia en la columna y tampoco fue posible recuperar "
                "un folio documental inequívoco desde Concepto."
            )
            st.dataframe(
                sin_ref_movs[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "concepto", "cargos", "abonos", "efecto_natural"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(refs_rec):
            st.markdown("#### 🟡 Referencias recuperadas desde Concepto")
            st.caption(
                "No se consideran 'sin referencia', pero se muestran para trazabilidad."
            )
            st.dataframe(
                refs_rec[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "concepto", "referencia_original", "referencia_norm",
                        "cargos", "abonos"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        reparadas = movs_vista[movs_vista["fila_reparada"]].copy()
        if len(reparadas):
            st.markdown("#### 🔧 Filas ARPON reconstruidas")
            st.caption(
                "El archivo fuente partió una partida en dos filas. El motor reconstruyó "
                "Concepto/Cargo/Abono/Saldo y posteriormente certificó el amarre completo."
            )
            st.dataframe(
                reparadas[[
                    "archivo", "fila_origen", "fila_continuacion", "fecha", "meta_codigo",
                    "poliza", "referencia_original", "concepto", "cargos", "abonos", "saldo_acumulado"
                ]],
                use_container_width=True, hide_index=True,
            )

        if len(negativos):
            st.markdown("#### 🟣 Montos negativos / reversos")
            st.caption(
                "Se señalan como movimiento especial; no se reinterpretan como abono."
            )
            st.dataframe(
                negativos[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "concepto", "referencia_norm", "cargos", "abonos",
                        "efecto_natural"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(duplicados):
            st.markdown("#### 🔁 Posibles duplicados exactos")
            st.caption(
                "Misma cuenta, fecha, tipo, póliza, concepto, referencia y monto. "
                "Es un indicador para revisión, no una conclusión automática."
            )
            st.dataframe(
                duplicados[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "tipo_poliza", "poliza", "concepto",
                        "referencia_norm", "cargos", "abonos"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(contrarios):
            st.markdown("#### ⚠️ Documentos con saldo contrario a la naturaleza")
            st.dataframe(
                contrarios[
                    [
                        "archivo", "meta_codigo", "meta_nombre", "naturaleza",
                        "referencia_norm", "primera_fecha", "dias",
                        "cargos", "abonos", "saldo_natural", "tipo_saldo"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if not any(
            [
                len(descuadres), len(sin_ref_movs), len(refs_rec), len(reparadas),
                len(negativos), len(duplicados), len(contrarios)
            ]
        ):
            st.success("Sin hallazgos relevantes con los criterios actuales.")

    # --------------------------------------------------------------------------
    # Semáforo
    # --------------------------------------------------------------------------
    with tabs[1]:
        st.subheader("🚦 Conciliación por cuenta")
        solo_problemas = st.toggle(
            "Ver solo cuentas con hallazgos",
            value=False,
            key="solo_problemas",
        )
        show = (
            audit_vista[audit_vista["estado"] != "🟢 OK"]
            if solo_problemas else audit_vista
        )

        cols = [
            "sistema_origen", "empresa", "tipo_cuenta", "archivo", "meta_codigo",
            "meta_nombre", "naturaleza", "naturaleza_confianza", "estado",
            "saldo_inicial", "saldo_inicial_pendiente",
            "total_cargos", "total_abonos", "saldo_final_aux", "saldo_final_pendiente",
            "movs_con_referencia", "movs_sin_referencia",
            "n_sin_referencia", "importe_bruto_sin_referencia",
            "n_refs_recuperadas", "n_referencias_libres",
            "n_montos_negativos", "descuadre_origen"
        ]
        st.dataframe(
            show[cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "saldo_inicial": st.column_config.NumberColumn(
                    "Saldo inicial", format="$%.2f"
                ),
                "saldo_inicial_pendiente": st.column_config.NumberColumn(
                    "Saldo inicial pendiente", format="$%.2f"
                ),
                "total_cargos": st.column_config.NumberColumn(
                    "Cargos", format="$%.2f"
                ),
                "total_abonos": st.column_config.NumberColumn(
                    "Abonos", format="$%.2f"
                ),
                "saldo_final_aux": st.column_config.NumberColumn(
                    "Saldo final ARPON", format="$%.2f"
                ),
                "saldo_final_pendiente": st.column_config.NumberColumn(
                    "Saldo pendiente", format="$%.2f"
                ),
                "movs_con_referencia": st.column_config.NumberColumn(
                    "Efecto con referencia", format="$%.2f"
                ),
                "movs_sin_referencia": st.column_config.NumberColumn(
                    "Efecto sin referencia", format="$%.2f"
                ),
                "importe_bruto_sin_referencia": st.column_config.NumberColumn(
                    "Bruto sin referencia", format="$%.2f"
                ),
                "descuadre_origen": st.column_config.NumberColumn(
                    "Descuadre", format="$%.2f"
                ),
            },
        )

    # --------------------------------------------------------------------------
    # Folios
    # --------------------------------------------------------------------------
    with tabs[2]:
        st.subheader("📑 Documentos / folios abiertos")
        st.caption(
            "Incluye folios prefijados, numéricos y documentos alfanuméricos reconocibles. "
            "La antigüedad es observada desde la primera fecha del folio, "
            "no fecha contractual de vencimiento."
        )

        orden = ["0-30", "31-60", "61-90", "90+"]
        positivos = folios_vista[folios_vista["saldo_natural"] > 0].copy()
        aging = (
            positivos.groupby("antiguedad_observada")["saldo_natural"]
            .agg(num_folios="count", saldo="sum")
            .reindex(orden)
            .fillna(0)
            .reset_index()
        )
        st.dataframe(aging, use_container_width=True, hide_index=True)

        if not folios_vista.empty:
            nat_sel = st.multiselect(
                "Naturaleza",
                sorted(folios_vista["naturaleza"].dropna().unique()),
                default=sorted(folios_vista["naturaleza"].dropna().unique()),
            )
            edades_sel = st.multiselect(
                "Antigüedad observada",
                orden,
                default=orden,
            )
            fv = folios_vista[
                folios_vista["naturaleza"].isin(nat_sel)
                & folios_vista["antiguedad_observada"].isin(edades_sel)
            ]
        else:
            fv = folios_vista

        st.dataframe(
            fv,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------------------------
    # Cruces / conciliación
    # --------------------------------------------------------------------------
    with tabs[3]:
        st.subheader("✅ Conciliación de partidas")
        c1, c2 = st.columns(2)
        c1.metric("Partidas conciliadas", f"{n_partidas_conciliadas:,}")
        c2.metric("Coincidencias a revisar", f"{n_partidas_revisar:,}")
        st.caption(
            "En pantalla y en el auxiliar descargado, verde significa grupo con neto aproximado "
            "a cero; amarillo significa efectos opuestos con remanente y requiere "
            "revisión."
        )

        st.markdown("#### Partidas marcadas")
        st.caption(
            "Estas son exactamente las filas que recibirán color y código en el "
            "auxiliar descargado."
        )
        partidas_base = movs_vista[movs_vista["conciliacion_marcada"]].copy()
        pc1, pc2 = st.columns(2)
        estados_disponibles = sorted(
            partidas_base["conciliacion_estado"].dropna().unique()
        )
        niveles_disponibles = sorted(
            partidas_base["conciliacion_nivel"].dropna()
            .loc[lambda s: s.astype(str).str.strip().ne("")].unique()
        )
        estados_conciliacion = pc1.multiselect(
            "Estado de conciliación",
            estados_disponibles,
            default=estados_disponibles,
            key="filtro_conciliacion_estado",
        )
        niveles_conciliacion = pc2.multiselect(
            "Nivel de evidencia",
            niveles_disponibles,
            default=niveles_disponibles,
            key="filtro_conciliacion_nivel",
        )
        partidas_pantalla = partidas_base[
            partidas_base["conciliacion_estado"].isin(estados_conciliacion)
            & partidas_base["conciliacion_nivel"].isin(niveles_conciliacion)
        ].copy()
        if partidas_pantalla.empty:
            st.info(
                "No hay partidas para marcar. Carga al mismo tiempo los auxiliares "
                "de las cuentas que deseas conciliar."
            )
        else:
            partidas_pantalla = partidas_pantalla[
                [
                    "conciliacion_estado", "conciliacion_nivel",
                    "conciliacion_criterio", "conciliacion_codigo",
                    "archivo", "fila_origen", "fecha", "meta_codigo",
                    "poliza", "referencia_original", "concepto",
                    "cargos", "abonos", "efecto_natural",
                ]
            ].sort_values(
                ["conciliacion_estado", "archivo", "fila_origen"]
            )
            partidas_pantalla = partidas_pantalla.rename(
                columns={
                    "conciliacion_estado": "Estado",
                    "conciliacion_nivel": "Nivel",
                    "conciliacion_criterio": "Criterio",
                    "conciliacion_codigo": "Código",
                    "archivo": "Archivo",
                    "fila_origen": "Fila ARPON",
                    "fecha": "Fecha",
                    "meta_codigo": "Cuenta",
                    "poliza": "Póliza",
                    "referencia_original": "Docto.",
                    "concepto": "Concepto",
                    "cargos": "Cargo",
                    "abonos": "Abono",
                    "efecto_natural": "Efecto natural",
                }
            )

            def color_partida(row):
                if row["Estado"] == "CONCILIADO":
                    estilo = "background-color: #EAF4E3; color: #006100;"
                else:
                    estilo = "background-color: #FFF7D6; color: #9C6500;"
                return [estilo] * len(row)

            tabla_marcada = (
                partidas_pantalla.style
                .apply(color_partida, axis=1)
                .format(
                    {
                        "Cargo": "${:,.2f}",
                        "Abono": "${:,.2f}",
                        "Efecto natural": "${:,.2f}",
                    },
                    na_rep="",
                )
            )
            st.dataframe(
                tabla_marcada,
                use_container_width=True,
                hide_index=True,
                height=min(620, 85 + 35 * len(partidas_pantalla)),
            )

        st.markdown("#### A. Cruces adicionales entre cuentas por el mismo documento")
        if cruces_ref_vista.empty:
            st.info(
                "No se encontraron documentos idénticos con efectos opuestos "
                "entre cuentas cargadas."
            )
        else:
            st.dataframe(
                cruces_ref_vista,
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("#### B. Coincidencias fuertes aunque el documento sea diferente")
        st.caption(
            "Misma fecha + mismo concepto + mismo importe absoluto + "
            "efecto pendiente opuesto entre cuentas del mismo tipo. Es evidencia para revisar/conciliar; "
            "no se basa en similitud difusa de nombres."
        )
        if evidencia_vista.empty:
            st.info(
                "No se encontraron coincidencias fuertes entre las cuentas cargadas."
            )
        else:
            st.success(
                f"Se encontraron {evidencia_vista['evidencia_id'].nunique():,} "
                "grupo(s) de evidencia."
            )
            st.dataframe(
                evidencia_vista,
                use_container_width=True,
                hide_index=True,
            )

    # --------------------------------------------------------------------------
    # Referencias
    # --------------------------------------------------------------------------
    with tabs[4]:
        st.subheader("🏷️ Auditoría de referencias")
        refs = tabla_referencias(movs_vista)

        tipos = (
            refs["referencia_tipo"].fillna("VACIA").value_counts()
            .rename_axis("tipo")
            .reset_index(name="movimientos")
        )
        st.dataframe(tipos, use_container_width=True, hide_index=True)

        filtro_tipo = st.multiselect(
            "Tipo de referencia",
            sorted(refs["referencia_tipo"].dropna().unique()),
            default=sorted(refs["referencia_tipo"].dropna().unique()),
        )
        refs_show = refs[refs["referencia_tipo"].isin(filtro_tipo)]
        st.dataframe(
            refs_show,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------------------------
    # Gráficos
    # --------------------------------------------------------------------------
    with tabs[5]:
        st.subheader("📉 Composición del saldo pendiente normalizado")

        saldo_ini = audit_vista["saldo_inicial_pendiente"].sum()
        con_ref = audit_vista["movs_con_referencia"].sum()
        sin_ref = audit_vista["movs_sin_referencia"].sum()
        desc = audit_vista["descuadre_origen"].sum()

        fig = go.Figure(
            data=[
                go.Bar(
                    name="Saldo inicial",
                    x=["Saldo total"],
                    y=[saldo_ini],
                ),
                go.Bar(
                    name="Efecto con referencia",
                    x=["Saldo total"],
                    y=[con_ref],
                ),
                go.Bar(
                    name="Efecto sin referencia",
                    x=["Saldo total"],
                    y=[sin_ref],
                ),
                go.Bar(
                    name="Descuadre",
                    x=["Saldo total"],
                    y=[desc],
                ),
            ]
        )
        fig.update_layout(
            barmode="relative",
            title="Composición del saldo pendiente normalizado",
            yaxis_title="Monto",
        )
        st.plotly_chart(fig, use_container_width=True)

    # --------------------------------------------------------------------------
    # Diagnóstico
    # --------------------------------------------------------------------------
    with tabs[6]:
        st.subheader("🧪 Diagnóstico técnico")

        st.markdown("#### Archivos")
        st.dataframe(diag_vista, use_container_width=True, hide_index=True)

        st.markdown("#### Detección de naturaleza")
        diag_cols = [
            "sistema_origen", "empresa", "tipo_cuenta", "archivo", "meta_codigo",
            "meta_nombre", "naturaleza", "naturaleza_confianza",
            "saldo_inicial", "total_cargos", "total_abonos", "esperado_arpon",
            "saldo_final_aux", "error_arpon", "saldo_final_pendiente",
            "ecuacion_saldo_fuente", "n_errores_saldo_secuencia",
            "max_error_saldo_secuencia", "error_ultimo_saldo_vs_total",
            "n_filas_reconstruidas"
        ]
        diag_cols = [c for c in diag_cols if c in df_audit.columns]
        st.dataframe(
            audit_vista[diag_cols],
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("#### Definiciones importantes")
        st.info(
            "• 'Sin referencia' significa Referencia/Docto. vacío y sin folio "
            "recuperable del Concepto.\n\n"
            "• 'Referencia libre' significa que sí existe texto en Referencia, "
            "pero no tiene forma de folio documental.\n\n"
            "• 'Antigüedad observada' no equivale a vencimiento.\n\n"
            "• 'Posible duplicado exacto' es un indicador, no una eliminación automática."
        )

    # --------------------------------------------------------------------------
    # Exportación completa
    # --------------------------------------------------------------------------
    st.divider()
    st.subheader("⬇️ Exportación")

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
    st.download_button(
        "⬇️ Descargar auditoría completa (Excel)",
        data=to_excel_workbook(export_tables),
        file_name="auditoria_master_saldos.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    if n_partidas_conciliadas or n_partidas_revisar or n_reparaciones:
        archivos_origen = [
            (uf.name, uf.getvalue()) for uf in uploaded_files
        ]
        data_marcada, nombre_marcado, mime_marcado = (
            construir_descarga_auxiliares_marcados(archivos_origen, movs)
        )
        st.download_button(
            "🎨 Descargar auxiliar(es) con conciliación marcada",
            data=data_marcada,
            file_name=nombre_marcado,
            mime=mime_marcado,
            help=(
                "Agrega una columna de auditoría y colorea conciliaciones/reconstrucciones "
                "sin modificar póliza, fecha, documento, concepto, cargos, abonos ni saldo del origen."
            ),
        )
    else:
        st.info(
            "No hay partidas de conciliación ni filas reconstruidas para marcar con los archivos cargados."
        )


if __name__ == "__main__":
    main()
